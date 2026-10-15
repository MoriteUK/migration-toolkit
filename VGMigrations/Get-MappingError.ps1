#Requires -Version 7.0
<#
.SYNOPSIS
    Stub — the Fly Graph API has no per-mapping error text endpoint.
.PARAMETER ProjectId
    GUID of the Fly project
.PARAMETER MappingId
    GUID of the specific mapping
#>
param(
    [Parameter(Mandatory=$true)][string]$ProjectId,
    [Parameter(Mandatory=$true)][string]$MappingId
)

# Error details are only accessible via async report generation in the Fly portal.
@{ Success = $false; ErrorMessage = '' } | ConvertTo-Json -Compress
