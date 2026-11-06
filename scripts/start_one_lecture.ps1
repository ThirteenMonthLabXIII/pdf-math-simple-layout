param([Parameter(Mandatory=$true)][string]$Pdf, [string]$WorkDir, [switch]$NoOcr)
$ErrorActionPreference='Stop'
$skill=Split-Path -Parent $PSScriptRoot
$python=Join-Path $skill '.venv\Scripts\python.exe'
if(!(Test-Path -LiteralPath $python)) {$python='python'}
$cli=Join-Path $PSScriptRoot 'layout_cli.py'
$source=(Resolve-Path -LiteralPath $Pdf).Path
if(!$WorkDir){$WorkDir=Join-Path (Split-Path -Parent $source) (([IO.Path]::GetFileNameWithoutExtension($source))+'_simple_layout')}
& $python $cli prepare --pdf $source --output $WorkDir
if($LASTEXITCODE -ne 0){throw 'PDF prepare failed'}
if(!$NoOcr){& $python $cli recognize --workspace $WorkDir; if($LASTEXITCODE -ne 0){throw 'OCR failed'}}
Write-Output ('Review and correct: '+(Join-Path $WorkDir 'layout.json'))
