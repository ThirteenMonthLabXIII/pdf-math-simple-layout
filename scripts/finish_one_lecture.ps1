param([Parameter(Mandatory=$true)][string]$WorkDir, [string]$Output)
$ErrorActionPreference='Stop'
$skill=Split-Path -Parent $PSScriptRoot
$python=Join-Path $skill '.venv\Scripts\python.exe'
if(!(Test-Path -LiteralPath $python)){$python='python'}
$cli=Join-Path $PSScriptRoot 'layout_cli.py'
$ws=(Resolve-Path -LiteralPath $WorkDir).Path
if(!$Output){$Output=Join-Path $ws 'candidate.docx'}
& $python $cli build --workspace $ws --profile (Join-Path $skill 'references\style_profile.json') --output $Output
if($LASTEXITCODE -ne 0){throw 'Word build failed'}
& $python $cli render --docx $Output --output (Join-Path $ws 'word_preview')
if($LASTEXITCODE -ne 0){throw 'Word render failed'}
& $python $cli audit --workspace $ws --docx $Output
if($LASTEXITCODE -ne 0){throw 'Word audit failed'}
Write-Output 'Review original and Word pages, then complete visual_review.json and audit again.'
