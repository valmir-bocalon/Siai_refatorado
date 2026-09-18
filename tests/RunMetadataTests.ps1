param([switch]$Negotiation, [switch]$Cobranca, [switch]$Diagnostic)
$ErrorActionPreference = 'Stop'
Set-Location (Split-Path $PSScriptRoot -Parent)
$mysqlBase = 'C:\Program Files (x86)\MySQL\MySQL Server 5.5'
$serverExe = Join-Path $mysqlBase 'bin\mysqld.exe'
$clientExe = Join-Path $mysqlBase 'bin\mysql.exe'
$adminExe = Join-Path $mysqlBase 'bin\mysqladmin.exe'
foreach ($file in @($serverExe, $clientExe, $adminExe, 'C:\Windows\SysWOW64\libmysql.dll')) {
    if (-not (Test-Path -LiteralPath $file)) { throw "Missing test dependency: $file" }
}
# Refuse to reuse any existing server, including one from a previous test.
$portCheck = [Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback, 33557)
try { $portCheck.Start() } finally { $portCheck.Stop() }
$runRoot = Join-Path (Get-Location).Path ('build\startup-audit\metadata-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path (Join-Path $runRoot 'data') -Force | Out-Null
$serverProcess = $null
$serverOwned = $false
try {
    $serverArgs = @('--no-defaults', ('--basedir="' + $mysqlBase + '"'),
        ('--datadir="' + $runRoot + '\data"'), '--bind-address=127.0.0.1', '--port=33557',
        '--skip-grant-tables', '--skip-name-resolve', '--skip-innodb', '--performance-schema=OFF',
        '--default-storage-engine=MyISAM', '--console', ('--pid-file="' + $runRoot + '\server.pid"'))
    $serverProcess = Start-Process -FilePath $serverExe -ArgumentList $serverArgs -WindowStyle Hidden -PassThru `
        -RedirectStandardOutput "$runRoot\stdout.log" -RedirectStandardError "$runRoot\stderr.log"
    for ($attempt = 0; $attempt -lt 40; $attempt++) {
        if ($serverProcess.HasExited) { throw "Disposable server exited. See $runRoot\stderr.log" }
        if (Test-Path -LiteralPath "$runRoot\server.pid") {
            $pidText = Get-Content -LiteralPath "$runRoot\server.pid" -Raw
            if ([int]$pidText -eq $serverProcess.Id) {
                $serverOwned = $true
                & $adminExe --no-defaults --host=127.0.0.1 --port=33557 --user=root ping 2>$null | Out-Null
                if ($LASTEXITCODE -eq 0) { break }
            }
        }
        Start-Sleep -Milliseconds 250
    }
    if (-not $serverOwned -or $attempt -eq 40) { throw 'Disposable server did not become ready' }
    # Only fictitious data; the application INI and production database are never read.
    $fixtureSQL = @'
CREATE DATABASE siai_metadata_test;
USE siai_metadata_test;
CREATE TABLE sample (
 id INTEGER NOT NULL PRIMARY KEY, category VARCHAR(10), label VARCHAR(80),
 amount DECIMAL(12,2), counter BIGINT, real_value DOUBLE, day_value DATE,
 time_value DATETIME, note TEXT, image BLOB, enabled TINYINT, flags BIT(3));
INSERT INTO sample VALUES
 (1,'a','first',12.34,1234567890123,1.5,'2026-09-16','2026-09-16 12:00:00','text','bytes',1,3),
 (2,'a','second',56.78,2234567890123,2.5,'2026-09-17','2026-09-17 13:00:00','text2','bytes2',0,1),
 (3,'b',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL);
'@
    $fixtureSQL | & $clientExe --no-defaults --host=127.0.0.1 --port=33557 --user=root --batch
    if ($LASTEXITCODE -ne 0) { throw 'Fixture creation failed' }
    $env:BDS = 'C:\Program Files (x86)\Embarcadero\Studio\18.0'
    $env:BDSINCLUDE = "$env:BDS\include"
    $env:BDSCOMMONDIR = 'C:\Users\Public\Documents\Embarcadero\Studio\18.0'
    $env:FrameworkDir = 'C:\Windows\Microsoft.NET\Framework\v3.5'
    $env:FrameworkVersion = 'v3.5'
    $env:LANGDIR = 'EN'
    $testName = if ($Cobranca) { 'CobrancaPerformanceSmoke' } elseif ($Negotiation) { 'NegotiationSmoke' } else { 'MetadataSmoke' }
    & 'C:\Windows\Microsoft.NET\Framework\v4.0.30319\MSBuild.exe' SIAI.dproj /t:Build `
        /p:Config=Release /p:Platform=Win32 "/p:MainSource=tests\$testName.dpr" `
        /p:DCC_ForceExecute=true /p:DCC_ExeOutput=.\build\startup-audit\metadata-tests `
        /p:DCC_DcuOutput=.\build\startup-audit\metadata-tests\dcu /v:minimal /nologo *> "$runRoot\build.log"
    if ($LASTEXITCODE -ne 0) { throw "Compilation failed: $runRoot\build.log" }
    if ($Diagnostic) {
        & ".\build\startup-audit\metadata-tests\$testName.exe" /diagnostico > "$runRoot\tests.log"
    } else {
        & ".\build\startup-audit\metadata-tests\$testName.exe" > "$runRoot\tests.log"
    }
    $testResult = $LASTEXITCODE
    Get-Content -LiteralPath "$runRoot\tests.log"
    if ($testResult -ne 0) { throw 'Metadata tests failed' }
    Write-Output "Logs: $runRoot"
} finally {
    if ($serverOwned -and -not $serverProcess.HasExited) {
        & $adminExe --no-defaults --host=127.0.0.1 --port=33557 --user=root shutdown
        if (-not $serverProcess.WaitForExit(10000)) {
            throw "Disposable server did not stop: PID $($serverProcess.Id)"
        }
    } elseif ($null -ne $serverProcess -and -not $serverProcess.HasExited) {
        Stop-Process -InputObject $serverProcess
    }
}
