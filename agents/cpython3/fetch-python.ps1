# Windows twin of fetch-python.sh: downloads the pinned embeddable CPython into agents/cpython3/python-embed/ (git-ignored).
$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$cfg = @{}
Get-Content (Join-Path $here "python-embed.txt") | Where-Object { $_ -match '^\w+=' } | ForEach-Object { $k, $v = $_ -split '=', 2; $cfg[$k] = $v.Trim() }
$dest = Join-Path $here "python-embed"
if (Test-Path (Join-Path $dest "python312.dll")) { exit 0 }
$url = "https://www.python.org/ftp/python/$($cfg.PY_VERSION)/python-$($cfg.PY_VERSION)-embed-amd64.zip"
$zip = Join-Path ([IO.Path]::GetTempPath()) "python-embed-$([guid]::NewGuid()).zip"
Write-Output "fetching $url"
Invoke-WebRequest -Uri $url -OutFile $zip -UseBasicParsing
$hash = (Get-FileHash $zip -Algorithm SHA256).Hash.ToLower()
if ($hash -ne $cfg.PY_SHA256) { Remove-Item $zip; throw "SHA-256 mismatch for $url (got $hash)" }
New-Item -ItemType Directory -Force $dest | Out-Null
Expand-Archive -Path $zip -DestinationPath $dest -Force
Remove-Item $zip
Remove-Item (Join-Path $dest "python.exe"), (Join-Path $dest "pythonw.exe") -ErrorAction SilentlyContinue
Write-Output "python-embed ready: $((Get-ChildItem $dest).Count) files"
