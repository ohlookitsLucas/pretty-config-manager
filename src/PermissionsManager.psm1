<#
  PermissionsManager.psm1
  Light-weight helpers to enumerate signed-in accounts and provide placeholders
  for setting calendar/mailbox permissions from a non-admin user context.
#>

function Get-SignedInAccounts {
    <#
      Returns a list of accounts from the local Outlook profile(s).
      Each item: @{ Name = 'Display Name'; SmtpAddress = 'user@example.com' }
    #>
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $ns = $outlook.GetNameSpace('MAPI')
        $accounts = @()
        if ($ns -and $ns.Accounts) {
            for ($i = 1; $i -le $ns.Accounts.Count; $i++) {
                $acc = $ns.Accounts.Item($i)
                $smtp = $null
                try { $smtp = $acc.SmtpAddress } catch { $smtp = '' }
                $accounts += [PSCustomObject]@{ Name = $acc.DisplayName; SmtpAddress = $smtp }
            }
        }
        return $accounts
    } catch {
        Write-Verbose "Outlook accounts could not be enumerated: $_"
        return @()
    }
}

function Set-CalendarPermission {
    param(
        [Parameter(Mandatory=$true)][string]$User,
        [Parameter(Mandatory=$true)][string]$Level,
        [switch]$Confirm
    )

    # This is a placeholder. Setting calendar permissions from a user-level tool may
    # require Exchange-level calls or using the Outlook object model with sufficient
    # server-side rights. For now, we provide a best-effort informational mapping
    # and return a not-implemented error to avoid silently failing changes.

    $mapping = @{
        'Can read' = 'Reviewer'
        'Can read & move/delete' = 'PublishingEditor'
        'Can reply' = 'Author'
    }

    $mapped = $null
    if ($mapping.ContainsKey($Level)) { $mapped = $mapping[$Level] } else { $mapped = $Level }

    Write-Verbose "Requested set calendar permission: User=$User Level=$Level Mapped=$mapped"
    throw "Set-CalendarPermission is not yet implemented in this build. This action requires Exchange or delegated mailbox-owner operations."
}

Export-ModuleMember -Function Get-SignedInAccounts,Set-CalendarPermission
