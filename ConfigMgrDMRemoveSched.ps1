#Requires -Version 5

<#
    .SYNOPSIS
    Identifies and modifies ConfigMgr Collections with Schedules and without Queries.

    .DESCRIPTION
    This function will loop over Configuration Manager collections to identify which have refresh schedules
    but do not have any queries.  Run with -WhatIf to query only.

    .EXAMPLE
    With no Parameters the function will loop over and set to Manual Refresh:
    .\ConfigMgrDMRemoveSched.ps1
    
    .EXAMPLE
    The OutputCSV allows you to output the collection information to a CSV file.
    .\ConfigMgrDMRemoveSched.ps1 -outputcsv $env:temp\Collections.csv

    .PARAMETER showgrid
    When the script is complete, output the results to a powershell grid.
    
    .PARAMETER outputcsv
    When the script is complete, write the output to the defined CSV file.

    .PARAMETER Force
    Don't prompt to continue when changing the RefreshType.

#>

### *** RUN AT YOUR OWN RISK AND ALWAYS TRY IN YOUR TEST ENVIRONMENT FIRST *** ###

<#
    Credits:
        Written by Chris Kibble (www.ChristopherKibble.com)

    With Assistance From:
        https://blogs.technet.microsoft.com/leesteve/2017/08/22/sccm-for-those-nasty-incremental-collections/#comment-3335
#>

[CmdletBinding(SupportsShouldProcess=$True)]
param(
    [switch]$showgrid,
    [string]$outputcsv = "",
    [switch]$Force
)

begin {
    Write-Verbose "Attempting to Import Configuration Manager Modules"
    Import-Module "$env:SMS_ADMIN_UI_PATH\..\ConfigurationManager.psd1" -ErrorAction Stop
    Try {
        $startPath = Get-Location
        set-location "$($(Get-PSDrive | Where-Object { $_.Provider.Name -eq "CMSITE" })[0].Name):"
    } Catch {
        Write-Error "Unable to connect to CMSite PSDrive"
        break
    }
}

process {
    
    $collections = @()

    Get-CMDeviceCollection | Where-Object { $_.RefreshType -ne 1 } | ForEach-Object {
        Write-Verbose "Checking Collection $($_.Name)"
        
        $collectionRules = $_.CollectionRules
        $queryCount = $($collectionRules | Where-Object { $_.SmsProviderObjectPath -eq "SMS_CollectionRuleQuery" }).Count
        
        $o = New-Object -TypeName PSObject
        Add-Member -InputObject $o -MemberType NoteProperty -Name "CollectionID" -Value $_.CollectionID
        Add-Member -InputObject $o -MemberType NoteProperty -Name "Name" -Value $_.Name
        Add-Member -InputObject $o -MemberType NoteProperty -Name "OwnedByThisSite" -Value $_.OwnedByThisSite
        Add-Member -InputObject $o -MemberType NoteProperty -Name "RefreshType" -Value $_.RefreshType
        Add-Member -InputObject $o -MemberType NoteProperty -Name "RefreshTypeNeedsModify" -Value $false
        Add-Member -InputObject $o -MemberType NoteProperty -Name "RefreshTypeModified" -Value $false
        Add-Member -InputObject $o -MemberType NoteProperty -Name "SMS_CollectionRuleDirect Count" -Value 0
        Add-Member -InputObject $o -MemberType NoteProperty -Name "SMS_CollectionRuleQuery Count" -Value 0
        Add-Member -InputObject $o -MemberType NoteProperty -Name "SMS_CollectionRuleIncludeCollection Count" -Value 0
        Add-Member -InputObject $o -MemberType NoteProperty -Name "SMS_CollectionRuleExcludeCollection Count" -Value 0

        $collectionrules | Group-Object -Property SmsProviderObjectPath | ForEach-Object {
            Add-Member -InputObject $o -MemberType NoteProperty -Name "$($_.Name) Count" -Value $_.Count -Force
        }

        if($queryCount -eq 0) {
            Out-Host -InputObject "Collection $($_.Name) ($($_.CollectionID)) has no queries but does have an update schedule."
            Add-Member -InputObject $o -MemberType NoteProperty -Name "RefreshTypeNeedsModify" -Value $true -Force
            If($_.OwnedByThisSite -eq $False) {
                Out-Host -InputObject "... Collection not owned by this site.  Taking no action."
            } else {
                If ($PSCmdlet.ShouldProcess($_.Name)) {
                    Write-Output "... Setting RefreshType of Collection to None"
                    If($Force -or $PScmdLet.ShouldContinue("Do you want to adjust the RefreshType on $($_.Name)?",$_.Name)) {
                        Set-CMCollection -CollectionId $($_.CollectionID) -RefreshType Manual
                        Add-Member -InputObject $o -MemberType NoteProperty -Name "RefreshTypeModified" -Value $true -Force
                    }                  
                }
            }
        }
        $collections += $o
    }

    Set-Location $startPath
    
    if($outputcsv) { $collections | Export-Csv $outputcsv -NoTypeInformation -whatif:$false }
    if($showgrid) { $collections | Out-GridView }

}


##########################################################################################################################################
#                                                                                                                                        #
# To Do Someday                                                                                                                          #
#                                                                                                                                        #
# * OutputCSV File Validation                                                                                                            #
# * Output to Log File in parallel with screen                                                                                           #
# * Write better function help
#                                                                                                                                        #
##########################################################################################################################################

