param([string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
$source = $PSScriptRoot
$destination = Join-Path $source '..\build'
if ($OutputDirectory) { $destination = $OutputDirectory }
New-Item -ItemType Directory -Path $destination -Force | Out-Null
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (-not (Test-Path -LiteralPath $compiler)) { $compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe' }
$references = @('/r:System.dll','/r:System.Core.dll','/r:Microsoft.CSharp.dll','/r:System.Drawing.dll','/r:System.Windows.Forms.dll','/r:System.Xml.dll','/r:System.Xml.Linq.dll','/r:System.IO.Compression.dll','/r:System.IO.Compression.FileSystem.dll')
$dll = Join-Path $destination 'ImageLink.UserAddIn.dll'
& $compiler /nologo /codepage:65001 /target:library /platform:anycpu /optimize+ "/out:$dll" @references (Join-Path $source 'Interop.cs') (Join-Path $source 'Engine.cs') (Join-Path $source 'AddIn.cs')
if ($LASTEXITCODE -ne 0) { throw 'Add-in compilation failed.' }
$setup = Join-Path $destination 'ImageLink-Setup.exe'
$verify32 = Join-Path $destination 'ImageLink-Verify-x86.exe'
$verify64 = Join-Path $destination 'ImageLink-Verify-x64.exe'
& $compiler /nologo /target:winexe /platform:x86 "/out:$verify32" /r:System.Xml.Linq.dll /r:System.Xml.dll (Join-Path $source 'VerifyInstall.cs')
if ($LASTEXITCODE -ne 0) { throw '32-bit verification compilation failed.' }
& $compiler /nologo /target:winexe /platform:x64 "/out:$verify64" /r:System.Xml.Linq.dll /r:System.Xml.dll (Join-Path $source 'VerifyInstall.cs')
if ($LASTEXITCODE -ne 0) { throw '64-bit verification compilation failed.' }
& $compiler /nologo /codepage:65001 /target:winexe /platform:anycpu /optimize+ "/out:$setup" "/resource:$dll,ImageLink.UserAddIn.dll" "/resource:$verify32,ImageLink-Verify-x86.exe" "/resource:$verify64,ImageLink-Verify-x64.exe" "/win32manifest:$(Join-Path $source 'app.manifest')" /r:System.dll /r:System.Core.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll (Join-Path $source 'Setup.cs')
if ($LASTEXITCODE -ne 0) { throw 'Installer compilation failed.' }
Write-Output "Built: $setup"


