
<#
    Version 1.00 [Author: Larry Song; Time: 2015-01-30]
        Add version logs
        Update function 'Add-Log' to add host colors
#>

function Add-Log
{
    PARAM(
        [String]$Path,
        [String]$Value,
        [String]$Type = 'Info'
    )
    $Type = $Type.ToUpper()
    $Date = Get-Date
    Write-Host "$($Date.ToString('[HH:mm:ss] '))[$Type] $Value" -ForegroundColor $(
        switch($Type)
        {
            'WARNING' {'Yellow'}
            'Error' {'Red'}
            default {'White'}
        }
    )
    if($Path){
        Add-Content -LiteralPath $Path -Value "$($Date.ToString('[HH:mm:ss] '))[$Type] $Value" -ErrorAction:SilentlyContinue
    }
}

function Get-AvailableVDIName
{
    PARAM(
        [parameter(Mandatory=$true)]
        [string]$User,

        [parameter(Mandatory=$true)]
        [string]$Prefix
    )
    $A_Int = [int][char]'A' - 1
    $User = $User -ireplace '^a', ''
    $VDINameAll = 1..26 | %{
        "$Prefix$([string][char]($_ + $A_Int))$User"
    }

    $objDomain = New-Object System.DirectoryServices.DirectoryEntry
    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher
    $objSearcher.SearchRoot = $objDomain
    $objSearcher.PageSize = 1000
    $objSearcher.Filter = "(&(objectCategory=Computer)(CN=${Prefix}*${User}))"
    $objSearcher.SearchScope = "Subtree"
    $objResults = $objSearcher.FindAll()
    $objResults = $objResults | %{$_.Properties["cn"]}

    $VDINameAll = $VDINameAll -ireplace "$($objResults -join '|')", ''
    return @($VDINameAll | ?{$_})
}

function Get-ADExisting
{
    PARAM(
        [parameter(Mandatory=$true)]
        [string]$SAMAccountName,

        [parameter(Mandatory=$true)]
        [string]$Type
    )

    switch ($Type){
    'computer' {
        $SAMAccountName = $SAMAccountName -ireplace '([^\$])$','$1$'
        break
    }
    'user' {
        break
    }
    default {
        Write-Error -Exception "Unknown type '$Type', type must be 'computer' or 'user'."
        return
    }
    }

    $objDomain = New-Object System.DirectoryServices.DirectoryEntry
    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher
    $objSearcher.SearchRoot = $objDomain
    $objSearcher.PageSize = 1000
    $objSearcher.Filter = "(&(objectCategory=$Type)(SAMAccountName=$SAMAccountName))"
    $objSearcher.SearchScope = "Subtree"
    return $objSearcher.FindAll()
}

function Get-VADObject
{
    PARAM(
        [parameter(Mandatory=$true)]
        [string]$SAMAccountName,

        [parameter(Mandatory=$true)]
        [string]$Type
    )

    switch ($Type){
    'computer' {
        $SAMAccountName = $SAMAccountName -ireplace '([^\$])$','$1$'
        break
    }
    'user' {
        break
    }
    default {
        Write-Error -Exception "Unknown type '$Type', type must be 'computer' or 'user'."
        return
    }
    }

    $objDomain = New-Object System.DirectoryServices.DirectoryEntry
    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher
    $objSearcher.SearchRoot = $objDomain
    $objSearcher.PageSize = 1000
    $objSearcher.Filter = "(&(objectCategory=$Type)(SAMAccountName=$SAMAccountName))"
    $objSearcher.SearchScope = "Subtree"
    return $objSearcher.FindAll()
}
