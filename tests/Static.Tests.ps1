Describe "WSUS Auditor credential handling" {
    It "does not pass a password to net use or a process argument" {
        $scriptText = Get-Content -LiteralPath (Join-Path (Get-Location) "WSUS_Auditor.ps1") -Raw -Encoding UTF8
        $scriptText | Should -Not -Match '(?im)net\s+use[^\r\n]*(password|txtPass|credential)'
        $scriptText | Should -Not -Match '(?im)ConnectionOptions[^\r\n]*|\.Password\s*=\s*\$password'
        $scriptText | Should -Match 'Get-Credential'
        $scriptText | Should -Match 'New-CimSession[^\r\n]*-Credential\s+\$credential'
    }
}
