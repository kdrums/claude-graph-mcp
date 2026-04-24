#Requires -Version 5.1
<#
.SYNOPSIS
    Builds and publishes McpServerGraphApi as a self-contained Windows x64 executable.

.DESCRIPTION
    Reads the version from McpServerGraphApi.csproj, runs dotnet publish, and places
    the output under artifacts\dist\<version>\ ready to copy to the Software Center source share.

.EXAMPLE
    .\build.ps1

.EXAMPLE
    .\build.ps1 -Verbose

.EXAMPLE
    .\build.ps1 -SignToolPath "C:\Program Files (x86)\Windows Kits\10\bin\x64\signtool.exe" -CodeSigningCertificateThumbprint "<thumbprint>" -RequireSignature
#>
[CmdletBinding()]
param(
    [string]$SignToolPath,
    [string]$CodeSigningCertificateThumbprint,
    [string]$TimestampUrl = 'http://timestamp.digicert.com',
    [switch]$RequireSignature
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$buildScript = Join-Path $PSScriptRoot 'eng\build.ps1'
if (-not (Test-Path $buildScript)) {
    Write-Error "Build script not found: $buildScript"
    exit 1
}

& $buildScript @PSBoundParameters
exit $LASTEXITCODE
