<#
  SignatureManager.Tests.ps1
  Pester 5 unit tests for SignatureManager.psm1
  Runs entirely offline — uses a temp sandbox for filesystem and registry.
#>

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\src\SignatureManager.psm1'
    Import-Module $modulePath -Force
}

Describe 'SignatureManager — Filesystem operations' {

    BeforeEach {
        # Set up a fresh sandbox per test
        $sandbox = & (Join-Path $PSScriptRoot 'Reset-TestEnvironment.ps1')
        Set-SignatureManagerPaths `
            -SignaturePath $sandbox.SignaturePath `
            -RegistryBase  $sandbox.RegBase `
            -ProfilesBase  $sandbox.ProfilesBase `
            -LogFile       $sandbox.LogFile `
            -BackupDir     $sandbox.BackupDir
    }

    AfterEach {
        & (Join-Path $PSScriptRoot 'Reset-TestEnvironment.ps1') -Teardown
    }

    Context 'Get-Signatures' {
        It 'Returns the 3 fixture signatures' {
            $sigs = Get-Signatures
            $sigs.Count | Should -Be 3
            $sigs | Should -Contain 'Corporate Standard'
            $sigs | Should -Contain 'Personal'
            $sigs | Should -Contain 'Shared Mailbox Sig'
        }

        It 'Returns empty array when folder is empty' {
            Get-ChildItem -Path $sandbox.SignaturePath -Recurse | Remove-Item -Recurse -Force
            $sigs = Get-Signatures
            $sigs | Should -HaveCount 0
        }
    }

    Context 'Get-SignatureStatus' {
        It 'Reports Corporate Standard as valid with _files folder' {
            $status = Get-SignatureStatus -Name 'Corporate Standard'
            $status.IsValid   | Should -BeTrue
            $status.HasHtm    | Should -BeTrue
            $status.HasTxt    | Should -BeTrue
            $status.HasFolder | Should -BeTrue
        }

        It 'Reports Personal as valid with missing .rtf warning' {
            $status = Get-SignatureStatus -Name 'Personal'
            $status.IsValid | Should -BeTrue
            $status.Warning | Should -BeLike '*no .rtf*'
        }

        It 'Reports nonexistent signature as invalid' {
            $status = Get-SignatureStatus -Name 'DoesNotExist'
            $status.IsValid | Should -BeFalse
        }
    }

    Context 'New-Signature' {
        It 'Creates a new signature with .htm and .txt files' {
            New-Signature -Name 'Test New Sig'
            Test-Path (Join-Path $sandbox.SignaturePath 'Test New Sig.htm') | Should -BeTrue
            Test-Path (Join-Path $sandbox.SignaturePath 'Test New Sig.txt') | Should -BeTrue
        }

        It 'Throws when name already exists' {
            { New-Signature -Name 'Corporate Standard' } | Should -Throw '*already exists*'
        }

        It 'Throws on invalid characters' {
            { New-Signature -Name 'Bad/Name' } | Should -Throw '*invalid characters*'
        }

        It 'Throws on blank name' {
            { New-Signature -Name '' } | Should -Throw
        }
    }

    Context 'Remove-Signature' {
        It 'Deletes all files for a signature' {
            Remove-Signature -Name 'Personal'
            Test-Path (Join-Path $sandbox.SignaturePath 'Personal.htm') | Should -BeFalse
            Test-Path (Join-Path $sandbox.SignaturePath 'Personal.txt') | Should -BeFalse
        }

        It 'Deletes _files folder when present' {
            Remove-Signature -Name 'Corporate Standard'
            Test-Path (Join-Path $sandbox.SignaturePath 'Corporate Standard_files') | Should -BeFalse
        }

        It 'Does not throw when deleting a signature without optional files' {
            { Remove-Signature -Name 'Personal' } | Should -Not -Throw
        }
    }

    Context 'Rename-Signature' {
        It 'Renames all files from old to new name' {
            Rename-Signature -OldName 'Personal' -NewName 'My Personal Sig'
            Test-Path (Join-Path $sandbox.SignaturePath 'Personal.htm')         | Should -BeFalse
            Test-Path (Join-Path $sandbox.SignaturePath 'My Personal Sig.htm')  | Should -BeTrue
            Test-Path (Join-Path $sandbox.SignaturePath 'My Personal Sig.txt')  | Should -BeTrue
        }

        It 'Renames _files folder when present' {
            Rename-Signature -OldName 'Corporate Standard' -NewName 'Rebranded'
            Test-Path (Join-Path $sandbox.SignaturePath 'Corporate Standard_files') | Should -BeFalse
            Test-Path (Join-Path $sandbox.SignaturePath 'Rebranded_files')          | Should -BeTrue
        }

        It 'Throws when target name already exists' {
            { Rename-Signature -OldName 'Personal' -NewName 'Corporate Standard' } | Should -Throw '*already exists*'
        }

        It 'No-ops when old and new names are identical' {
            { Rename-Signature -OldName 'Personal' -NewName 'Personal' } | Should -Not -Throw
            Test-Path (Join-Path $sandbox.SignaturePath 'Personal.htm') | Should -BeTrue
        }
    }

    Context 'Copy-Signature' {
        It 'Duplicates all files to a new name' {
            Copy-Signature -SourceName 'Personal' -TargetName 'Personal Copy'
            Test-Path (Join-Path $sandbox.SignaturePath 'Personal.htm')      | Should -BeTrue
            Test-Path (Join-Path $sandbox.SignaturePath 'Personal Copy.htm') | Should -BeTrue
            Test-Path (Join-Path $sandbox.SignaturePath 'Personal Copy.txt') | Should -BeTrue
        }

        It 'Duplicates _files folder when present' {
            Copy-Signature -SourceName 'Corporate Standard' -TargetName 'Corp Copy'
            Test-Path (Join-Path $sandbox.SignaturePath 'Corp Copy_files\logo.png') | Should -BeTrue
        }

        It 'Throws when target already exists' {
            { Copy-Signature -SourceName 'Personal' -TargetName 'Corporate Standard' } | Should -Throw '*already exists*'
        }

        It 'Throws when source does not exist' {
            { Copy-Signature -SourceName 'Nope' -TargetName 'Whatever' } | Should -Throw '*not found*'
        }
    }

    Context 'Save-SignatureHtml' {
        It 'Overwrites .htm and regenerates .txt' {
            $newHtml = '<html><body><p>Updated content</p></body></html>'
            Save-SignatureHtml -Name 'Personal' -HtmlContent $newHtml
            $htm = Get-Content (Join-Path $sandbox.SignaturePath 'Personal.htm') -Raw
            $htm | Should -BeLike '*Updated content*'
            $txt = Get-Content (Join-Path $sandbox.SignaturePath 'Personal.txt') -Raw
            $txt | Should -BeLike '*Updated content*'
        }
    }
}

Describe 'SignatureManager — Registry operations' {

    BeforeEach {
        $sandbox = & (Join-Path $PSScriptRoot 'Reset-TestEnvironment.ps1')
        Set-SignatureManagerPaths `
            -SignaturePath $sandbox.SignaturePath `
            -RegistryBase  $sandbox.RegBase `
            -ProfilesBase  $sandbox.ProfilesBase `
            -LogFile       $sandbox.LogFile `
            -BackupDir     $sandbox.BackupDir
    }

    AfterEach {
        & (Join-Path $PSScriptRoot 'Reset-TestEnvironment.ps1') -Teardown
    }

    Context 'Get-SignatureAssignments' {
        It 'Returns 4 entries (1 global + 3 per-account)' {
            $assignments = Get-SignatureAssignments
            $assignments.Count | Should -Be 4
        }

        It 'First entry is the global default' {
            $assignments = Get-SignatureAssignments
            $assignments[0].AccountName | Should -Be 'All accounts (global default)'
            $assignments[0].NewSignature | Should -Be 'Corporate Standard'
            $assignments[0].ReplySignature | Should -Be 'Personal'
        }

        It 'Contains alice@contoso.com with correct signatures' {
            $assignments = Get-SignatureAssignments
            $alice = $assignments | Where-Object { $_.SmtpAddress -eq 'alice@contoso.com' }
            $alice | Should -Not -BeNullOrEmpty
            $alice.NewSignature   | Should -Be 'Corporate Standard'
            $alice.ReplySignature | Should -Be 'Personal'
        }

        It 'Contains shared@contoso.com' {
            $assignments = Get-SignatureAssignments
            $shared = $assignments | Where-Object { $_.SmtpAddress -eq 'shared@contoso.com' }
            $shared | Should -Not -BeNullOrEmpty
            $shared.NewSignature | Should -Be 'Shared Mailbox Sig'
        }

        It 'Contains bob.delegate@contoso.com with empty sigs' {
            $assignments = Get-SignatureAssignments
            $bob = $assignments | Where-Object { $_.SmtpAddress -eq 'bob.delegate@contoso.com' }
            $bob | Should -Not -BeNullOrEmpty
            $bob.NewSignature   | Should -Be ''
            $bob.ReplySignature | Should -Be ''
        }
    }

    Context 'Set-SignatureAssignment' {
        It 'Updates global NewSignature' {
            Set-SignatureAssignment -RegistryPath $sandbox.RegBase -NewSignature 'Personal'
            $val = (Get-ItemProperty -Path $sandbox.RegBase).NewSignature
            $val | Should -Be 'Personal'
        }

        It 'Creates a registry backup file' {
            Set-SignatureAssignment -RegistryPath $sandbox.RegBase -NewSignature 'Test'
            $backups = Get-ChildItem -Path $sandbox.BackupDir -Filter '*.reg' -ErrorAction SilentlyContinue
            $backups.Count | Should -BeGreaterThan 0
        }

        It 'Updates per-account signature' {
            $assignments = Get-SignatureAssignments
            $alice = $assignments | Where-Object { $_.SmtpAddress -eq 'alice@contoso.com' }
            Set-SignatureAssignment -RegistryPath $alice.RegistryPath -NewSignature 'Personal'
            $updated = Get-SignatureAssignments
            $aliceUpdated = $updated | Where-Object { $_.SmtpAddress -eq 'alice@contoso.com' }
            $aliceUpdated.NewSignature | Should -Be 'Personal'
        }
    }

    Context 'Rename updates registry references' {
        It 'Updates all accounts that reference the renamed signature' {
            Rename-Signature -OldName 'Corporate Standard' -NewName 'New Brand'
            $assignments = Get-SignatureAssignments
            $global = $assignments[0]
            $global.NewSignature | Should -Be 'New Brand'
            $alice = $assignments | Where-Object { $_.SmtpAddress -eq 'alice@contoso.com' }
            $alice.NewSignature | Should -Be 'New Brand'
        }
    }
}

Describe 'SignatureManager — Export/Import' {

    BeforeEach {
        $sandbox = & (Join-Path $PSScriptRoot 'Reset-TestEnvironment.ps1')
        Set-SignatureManagerPaths `
            -SignaturePath $sandbox.SignaturePath `
            -RegistryBase  $sandbox.RegBase `
            -ProfilesBase  $sandbox.ProfilesBase `
            -LogFile       $sandbox.LogFile `
            -BackupDir     $sandbox.BackupDir
    }

    AfterEach {
        & (Join-Path $PSScriptRoot 'Reset-TestEnvironment.ps1') -Teardown
    }

    It 'Export then Import round-trips a signature' {
        $zipPath = Join-Path $sandbox.SandboxRoot 'export-test.zip'
        Export-Signature -Name 'Corporate Standard' -Destination $zipPath
        Test-Path $zipPath | Should -BeTrue

        # Delete the original
        Remove-Signature -Name 'Corporate Standard'
        (Get-Signatures) | Should -Not -Contain 'Corporate Standard'

        # Re-import
        Import-Signature -ZipPath $zipPath
        (Get-Signatures) | Should -Contain 'Corporate Standard'
    }
}
