$repoRoot = Split-Path -Parent $PSScriptRoot
$powerShellFiles = @(
    Get-ChildItem -Path $repoRoot -Recurse -File |
        Where-Object {
            ($_.Extension -eq '.ps1' -or $_.Extension -eq '.psm1') -and
            $_.FullName -notmatch '[\\/]\.git[\\/]'
        }
)

Describe 'PowerShell syntax' {
    foreach ($file in $powerShellFiles) {
        $relativePath = $file.FullName.Substring($repoRoot.Length).TrimStart('\', '/')

        It "$relativePath parses without syntax errors" {
            $tokens = $null
            $parseErrors = $null

            [System.Management.Automation.Language.Parser]::ParseFile(
                $file.FullName,
                [ref]$tokens,
                [ref]$parseErrors
            ) | Out-Null

            if ($parseErrors.Count -gt 0) {
                $messages = @($parseErrors | ForEach-Object { $_.Message }) -join '; '
                Write-Host "Parse errors: $messages"
            }

            $parseErrors.Count | Should Be 0
        }
    }
}
