param([string]$OutputDirectory=(Join-Path $PSScriptRoot 'dist/겹서랍'))
$ErrorActionPreference='Stop'
if($PSVersionTable.PSEdition -ne 'Desktop'){throw 'Windows PowerShell 5.1의 powershell.exe로 build.ps1을 실행하세요.'}
$source=Join-Path $PSScriptRoot 'src'
$output=[IO.Path]::GetFullPath($OutputDirectory)
if(Test-Path -LiteralPath $output){throw '출력 폴더가 이미 있습니다. 새 OutputDirectory를 지정하세요.'}
[void][IO.Directory]::CreateDirectory($output)
Get-ChildItem -LiteralPath $source -File|Where-Object {$_.Extension -in @('.ps1','.cs')}|ForEach-Object {Copy-Item -LiteralPath $_.FullName -Destination (Join-Path $output $_.Name)}
Copy-Item -LiteralPath (Join-Path $source 'assets') -Destination (Join-Path $output 'assets') -Recurse
$compiler=Join-Path ([Runtime.InteropServices.RuntimeEnvironment]::GetRuntimeDirectory()) 'csc.exe'
if(!(Test-Path -LiteralPath $compiler)){throw '.NET Framework C# 컴파일러를 찾지 못했습니다.'}
$automation=[Management.Automation.PSObject].Assembly.Location
& $compiler /nologo /target:winexe /platform:x64 (('/out:')+(Join-Path $output 'START.exe')) (('/win32icon:')+(Join-Path $source 'assets/gyeopseorap-open.ico')) (('/reference:')+$automation) /reference:System.Windows.Forms.dll (Join-Path $source 'Launcher.cs')
if($LASTEXITCODE -ne 0){throw '빌드 실패'}
Write-Output ('빌드 완료: '+(Join-Path $output 'START.exe'))
