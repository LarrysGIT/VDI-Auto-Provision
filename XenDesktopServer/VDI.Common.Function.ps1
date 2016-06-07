
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
        [string]$Type,

        [parameter()]
        [string[]]$Properties
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
    if($Properties){
        $Properties | %{$objSearcher.PropertiesToLoad.Add($_)} | Out-Null
    }
    return $objSearcher.FindAll()
}

function HashConv
{
    PARAM(
        $strHash
    )
	return [System.Runtime.InteropServices.Marshal]::PtrToStringUni(
        [System.Runtime.InteropServices.Marshal]::SecureStringToCoTaskMemUnicode($(ConvertTo-SecureString -String $strHash))
    )
}

function FillZero
{
    PARAM(
        [string]$InStr,
        [int]$Len
    )
    if($InStr.Length -ge $Len){return $InStr}
    return "$("0" * ($Len - $InStr.Length))$InStr"
}

function TemplatePick
{
    PARAM(
        [string]$Cluster,
        [string]$Template,
        [string]$CurrentMax,
        [switch]$Auto
    )
    switch -regex ($CurrentMax){
    '^(.+?)(\d+)$' {
        return $Matches[1], ([int]($Matches[2]) + 1), ($Matches[2]).Length
        break
    }
    default {
        if($Auto){
            TemplatePick -Cluster $Cluster -CurrentMax $(
                @(Get-Cluster $Cluster | Get-VM | %{
                    if($_.ExtensionData.Config.Files.VmPathName -imatch "($Template.*?)/")
                    {
                        $Matches[1]
                    }
                } | Sort-Object)[-1].Name
            ) -Template $Template
        }else{
            return $Template, "1", 3
        }
    }
    }
}
