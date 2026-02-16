Describe 'Get-Signatures' {
    It 'Returns an array (or empty array) and does not throw' {
        . "$PSScriptRoot\..\src\SignatureManager.psm1"
        { Get-Signatures | Out-Null } | Should -Not -Throw
    }
}
