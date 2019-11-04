# Stop execution when an error occurs
$ErrorActionPreference = "Stop"

# Contains the path of regasm.exe
$RegAsmPath = "$($env:Windir)\Microsoft.NET\Framework\v2.0.50727\regasm.exe"
if (-not (Test-Path $RegAsmPath))
{
    throw "regasm.exe not found in $RegAsmPath"
}

# Ask for [Install] or [Abort]
$Options = [System.Management.Automation.Host.ChoiceDescription[]] @("&Install", "&Abort")
$Selection = $host.UI.PromptForChoice("OutlookCOMM for NAV Classic installer", "What do you want to do?", $Options, 0)
if ($Selection -eq 0)
{
    # Create the folder OutlookCOMM in C:\ drive if it does not exists
    $TargetFolder = "C:\OutlookCOMM"
    if (-not [IO.Directory]::Exists($TargetFolder))
    {
        [IO.Directory]::CreateDirectory($TargetFolder) | Out-Null
    }

    # Copy OutlookCOMM files to $TargetFolder
    Copy-Item "$PSScriptRoot\OutlookCOMM.COM.dll" -Destination "$TargetFolder" -Force
    Copy-Item "$PSScriptRoot\OutlookCOMM.COM.tlb" -Destination "$TargetFolder" -Force
    
    # Register the COM component with regasm.exe
    $RegAsmArgs = ($("`"$TargetFolder\OutlookCOMM.COM.dll`""), "/tlb:`"$("$TargetFolder\OutlookCOMM.COM.tlb")`"", "/silent", "/nologo", "/codebase")  
    & $RegAsmPath @RegAsmArgs
}
else
{
    # Exit if [Abort]
    exit(0)
}