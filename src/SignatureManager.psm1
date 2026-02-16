<#
  SignatureManager.psm1
  Functions to list, copy, export and import Outlook signatures for the current Windows user.
#>

function Get-Signatures {
    $sigPath = Join-Path $env:APPDATA 'Microsoft\Signatures'
    if (-not (Test-Path $sigPath)) { return @() }
    Get-ChildItem -Path $sigPath -Filter *.htm -File -ErrorAction SilentlyContinue | ForEach-Object { $_.BaseName } | Sort-Object -Unique
}

function Get-SignatureHtmlPath {
    param([Parameter(Mandatory=$true)][string]$Name)
    $sigPath = Join-Path $env:APPDATA 'Microsoft\Signatures'
    return Join-Path $sigPath ($Name + '.htm')
}

function Copy-Signature {
    param(
        [Parameter(Mandatory=$true)][string]$SourceName,
        [Parameter(Mandatory=$true)][string]$TargetName
    )

    $sigPath = Join-Path $env:APPDATA 'Microsoft\Signatures'
    if (-not (Test-Path $sigPath)) { throw "Signatures folder not found: $sigPath" }

    $files = @("$SourceName.htm","$SourceName.rtf","$SourceName.txt")
    foreach ($f in $files) {
        $src = Join-Path $sigPath $f
        if (Test-Path $src) {
            $dest = Join-Path $sigPath ($f -replace [regex]::Escape($SourceName), $TargetName)
            Copy-Item -Path $src -Destination $dest -Force
        }
    }

    # copy resources folder if present (SourceName_files)
    $srcFolder = Join-Path $sigPath ($SourceName + '_files')
    if (Test-Path $srcFolder) {
        $destFolder = Join-Path $sigPath ($TargetName + '_files')
        if (Test-Path $destFolder) { Remove-Item -Path $destFolder -Recurse -Force }
        Copy-Item -Path $srcFolder -Destination $destFolder -Recurse -Force
    }
}

function Export-Signature {
    param(
        [Parameter(Mandatory=$true)][string]$Name,
        [Parameter(Mandatory=$true)][string]$Destination
    )
    $sigPath = Join-Path $env:APPDATA 'Microsoft\Signatures'
    if (-not (Test-Path $sigPath)) { throw "Signatures folder not found: $sigPath" }

    $temp = Join-Path ([System.IO.Path]::GetTempPath()) ([System.IO.Path]::GetRandomFileName())
    New-Item -ItemType Directory -Path $temp | Out-Null
    try {
        $patterns = @("$Name.*", "$Name`_files")
        foreach ($p in $patterns) {
            Get-ChildItem -Path $sigPath -Filter $p -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
                $rel = $_.FullName.Substring($sigPath.Length).TrimStart('\')
                $dest = Join-Path $temp $rel
                $destDir = Split-Path $dest -Parent
                if (-not (Test-Path $destDir)) { New-Item -ItemType Directory -Path $destDir | Out-Null }
                Copy-Item -Path $_.FullName -Destination $dest -Force
            }
        }

        # compress
        if (Test-Path $Destination) { Remove-Item -Path $Destination -Force }
        Compress-Archive -Path (Join-Path $temp '*') -DestinationPath $Destination -Force
    } finally {
        Remove-Item -Path $temp -Recurse -Force
    }
}

function Import-Signature {
    param(
        [Parameter(Mandatory=$true)][string]$ZipPath
    )
    if (-not (Test-Path $ZipPath)) { throw "Zip not found: $ZipPath" }
    $sigPath = Join-Path $env:APPDATA 'Microsoft\Signatures'
    $temp = Join-Path ([System.IO.Path]::GetTempPath()) ([System.IO.Path]::GetRandomFileName())
    New-Item -ItemType Directory -Path $temp | Out-Null
    try {
        Expand-Archive -Path $ZipPath -DestinationPath $temp -Force
        Get-ChildItem -Path $temp -Recurse | ForEach-Object {
            $rel = $_.FullName.Substring($temp.Length).TrimStart('\')
            $dest = Join-Path $sigPath $rel
            $destDir = Split-Path $dest -Parent
            if (-not (Test-Path $destDir)) { New-Item -ItemType Directory -Path $destDir | Out-Null }
            Copy-Item -Path $_.FullName -Destination $dest -Force
        }
    } finally {
        Remove-Item -Path $temp -Recurse -Force
    }
}

Export-ModuleMember -Function Get-Signatures,Get-SignatureHtmlPath,Copy-Signature,Export-Signature,Import-Signature
