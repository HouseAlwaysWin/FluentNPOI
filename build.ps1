# FluentNPOI Build Script
# 本地構建和測試腳本

param(
    [Parameter(Mandatory=$false)]
    [ValidateSet('Build', 'Test', 'Pack', 'Clean', 'All')]
    [string]$Task = 'All',
    
    [Parameter(Mandatory=$false)]
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Release',
    
    [Parameter(Mandatory=$false)]
    [string]$Version = '1.0.0'
)

$ErrorActionPreference = 'Stop'
$ProjectRoot = $PSScriptRoot
$ProjectFile = Join-Path $ProjectRoot "FluentNPOI\FluentNPOI.csproj"
$TestProject = Join-Path $ProjectRoot "FluentNPOIUnitTest\FluentNPOIUnitTest.csproj"
$OutputDir = Join-Path $ProjectRoot "artifacts"

function Write-TaskHeader {
    param([string]$Message)
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host $Message -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Cyan
}

function Clean-Build {
    Write-TaskHeader "清理構建產物 / Cleaning build artifacts"
    
    if (Test-Path $OutputDir) {
        Remove-Item -Path $OutputDir -Recurse -Force
        Write-Host "已刪除 artifacts 目錄" -ForegroundColor Green
    }
    
    dotnet clean $ProjectFile --configuration $Configuration
    dotnet clean $TestProject --configuration $Configuration
    
    Write-Host "清理完成 / Clean completed" -ForegroundColor Green
}

function Build-Project {
    Write-TaskHeader "構建專案 / Building projects"
    
    # 定義所有需要構建的專案 (按依賴順序排列)
    $BuildProjects = @(
        @{ Name = "FluentNPOI";           Path = Join-Path $ProjectRoot "FluentNPOI\FluentNPOI.csproj" },
        @{ Name = "FluentNPOI.Streaming"; Path = Join-Path $ProjectRoot "FluentNPOI.Streaming\FluentNPOI.Streaming.csproj" },
        @{ Name = "FluentNPOI.Pdf";       Path = Join-Path $ProjectRoot "FluentNPOI.Pdf\FluentNPOI.Pdf.csproj" },
        @{ Name = "FluentNPOI.Charts";    Path = Join-Path $ProjectRoot "FluentNPOI.Charts\FluentNPOI.Charts.csproj" },
        @{ Name = "FluentNPOI.All";       Path = Join-Path $ProjectRoot "FluentNPOI.All\FluentNPOI.All.csproj" }
    )
    
    Write-Host "恢復依賴 / Restoring dependencies..." -ForegroundColor Yellow
    dotnet restore (Join-Path $ProjectRoot "FluentNPOI.sln")
    
    foreach ($Project in $BuildProjects) {
        Write-Host "`n構建 $($Project.Name) / Building $($Project.Name)..." -ForegroundColor Yellow
        dotnet build $Project.Path --configuration $Configuration --no-restore
        
        if ($LASTEXITCODE -ne 0) {
            throw "$($Project.Name) 構建失敗 / Build failed"
        }
        Write-Host "  ✅ $($Project.Name) 構建成功 / Build succeeded" -ForegroundColor Green
    }
    
    Write-Host "`n所有專案構建完成 / All projects built successfully" -ForegroundColor Green
}

function Test-Project {
    Write-TaskHeader "運行測試 / Running tests"
    
    Write-Host "恢復測試專案依賴 / Restoring test project dependencies..." -ForegroundColor Yellow
    dotnet restore $TestProject
    
    Write-Host "`n構建測試專案 / Building test project..." -ForegroundColor Yellow
    dotnet build $TestProject --configuration $Configuration --no-restore
    
    if ($LASTEXITCODE -ne 0) {
        throw "測試專案構建失敗 / Test project build failed"
    }
    
    Write-Host "`n運行單元測試 / Running unit tests..." -ForegroundColor Yellow
    dotnet test $TestProject `
        --configuration $Configuration `
        --no-build `
        --verbosity normal `
        --logger "console;verbosity=detailed" `
        --collect:"XPlat Code Coverage" `
        --results-directory "./TestResults"
    
    if ($LASTEXITCODE -ne 0) {
        throw "測試失敗 / Tests failed"
    }
    
    Write-Host "`n所有測試通過 / All tests passed" -ForegroundColor Green
}

function Pack-Project {
    Write-TaskHeader "打包 NuGet 套件 / Packing NuGet packages"
    
    if (-not (Test-Path $OutputDir)) {
        New-Item -Path $OutputDir -ItemType Directory | Out-Null
    }
    
    Write-Host "版本號 / Version: $Version" -ForegroundColor Yellow
    
    # 定義所有需要打包的專案 (按依賴順序排列)
    $PackageProjects = @(
        @{ Name = "FluentNPOI";           Path = Join-Path $ProjectRoot "FluentNPOI\FluentNPOI.csproj" },
        @{ Name = "FluentNPOI.Streaming"; Path = Join-Path $ProjectRoot "FluentNPOI.Streaming\FluentNPOI.Streaming.csproj" },
        @{ Name = "FluentNPOI.Pdf";       Path = Join-Path $ProjectRoot "FluentNPOI.Pdf\FluentNPOI.Pdf.csproj" },
        @{ Name = "FluentNPOI.Charts";    Path = Join-Path $ProjectRoot "FluentNPOI.Charts\FluentNPOI.Charts.csproj" },
        @{ Name = "FluentNPOI.All";       Path = Join-Path $ProjectRoot "FluentNPOI.All\FluentNPOI.All.csproj" }
    )
    
    $SuccessCount = 0
    $FailedPackages = @()
    
    foreach ($Project in $PackageProjects) {
        Write-Host "`n打包 $($Project.Name) / Packing $($Project.Name)..." -ForegroundColor Yellow
        
        dotnet pack $Project.Path `
            --configuration $Configuration `
            --no-build `
            --output $OutputDir `
            /p:PackageVersion=$Version `
            /p:Version=$Version
        
        if ($LASTEXITCODE -ne 0) {
            Write-Host "  ❌ $($Project.Name) 打包失敗 / Pack failed" -ForegroundColor Red
            $FailedPackages += $Project.Name
        } else {
            Write-Host "  ✅ $($Project.Name) 打包成功 / Pack succeeded" -ForegroundColor Green
            $SuccessCount++
        }
    }
    
    if ($FailedPackages.Count -gt 0) {
        throw "以下套件打包失敗 / The following packages failed to pack: $($FailedPackages -join ', ')"
    }
    
    Write-Host "`n打包完成 / Pack completed ($SuccessCount packages)" -ForegroundColor Green
    Write-Host "輸出目錄 / Output directory: $OutputDir" -ForegroundColor Cyan
    Get-ChildItem -Path $OutputDir -Filter *.nupkg | ForEach-Object {
        Write-Host "  - $($_.Name)" -ForegroundColor White
    }
}

function Show-Coverage {
    Write-TaskHeader "生成覆蓋率報告 / Generating coverage report"
    
    $CoverageFiles = Get-ChildItem -Path "./TestResults" -Filter "coverage.cobertura.xml" -Recurse
    
    if ($CoverageFiles.Count -eq 0) {
        Write-Host "未找到覆蓋率文件，請先運行測試 / No coverage files found, please run tests first" -ForegroundColor Yellow
        return
    }
    
    Write-Host "安裝 ReportGenerator / Installing ReportGenerator..." -ForegroundColor Yellow
    dotnet tool install -g dotnet-reportgenerator-globaltool --ignore-failed-sources 2>$null
    
    $ReportDir = Join-Path $ProjectRoot "CoverageReport"
    
    Write-Host "`n生成報告 / Generating report..." -ForegroundColor Yellow
    reportgenerator `
        -reports:"./TestResults/**/coverage.cobertura.xml" `
        -targetdir:$ReportDir `
        -reporttypes:"Html;HtmlSummary"
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host "`n覆蓋率報告已生成 / Coverage report generated" -ForegroundColor Green
        Write-Host "報告位置 / Report location: $ReportDir" -ForegroundColor Cyan
        
        $IndexFile = Join-Path $ReportDir "index.html"
        if (Test-Path $IndexFile) {
            Write-Host "`n正在打開報告 / Opening report..." -ForegroundColor Yellow
            Start-Process $IndexFile
        }
    }
}

function Show-Info {
    Write-TaskHeader "環境資訊 / Environment Information"
    
    Write-Host ".NET SDK 版本 / .NET SDK Version:" -ForegroundColor Yellow
    dotnet --version
    
    Write-Host "`n.NET Runtimes:" -ForegroundColor Yellow
    dotnet --list-runtimes
    
    Write-Host "`n專案資訊 / Project Information:" -ForegroundColor Yellow
    Write-Host "  專案文件 / Project file: $ProjectFile" -ForegroundColor White
    Write-Host "  配置 / Configuration: $Configuration" -ForegroundColor White
    Write-Host "  版本 / Version: $Version" -ForegroundColor White
}

# 主程序 / Main
try {
    Show-Info
    
    switch ($Task) {
        'Clean' {
            Clean-Build
        }
        'Build' {
            Build-Project
        }
        'Test' {
            Build-Project
            Test-Project
            Show-Coverage
        }
        'Pack' {
            Build-Project
            Test-Project
            Pack-Project
        }
        'All' {
            Clean-Build
            Build-Project
            Test-Project
            Pack-Project
            Show-Coverage
        }
    }
    
    Write-Host "`n✅ 任務完成 / Task completed successfully" -ForegroundColor Green
    exit 0
}
catch {
    Write-Host "`n❌ 錯誤 / Error: $_" -ForegroundColor Red
    exit 1
}

