# Stop execution when an error occurs
$ErrorActionPreference = "Stop"

# Detect if PowerShell has administrator privileges
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
{   
	# Run the script with PowerShell with administrator privileges
    $Script = $MyInvocation.MyCommand.Definition
    $Ps = Join-Path $PSHome 'powershell.exe'
    Start-Process $Ps -Verb runas -ArgumentList "& '$Script'"
    exit(0)
}
else
{
    # Contains the path of regasm.exe
    $RegAsmPath = "$($env:Windir)\Microsoft.NET\Framework\v2.0.50727\regasm.exe"
    if (-not (Test-Path $RegAsmPath))
    {
        throw "regasm.exe not found in $RegAsmPath"
    }

    # Ask for [Install], [Remove] or [Abort]
    $Options = [System.Management.Automation.Host.ChoiceDescription[]] @("&Install", "&Remove", "&Abort")
    $Selection = $host.UI.PromptForChoice("Installer of OutlookCOMM for NAV Classic", "What do you want to do?", $Options, 0)
    if ($Selection -eq 0 -or $Selection -eq 1)
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
    
        # Create the [Install] command
        $RegAsmArgs = ($("`"$TargetFolder\OutlookCOMM.COM.dll`""), "/tlb:`"$("$TargetFolder\OutlookCOMM.COM.tlb")`"", "/silent", "/codebase")

        # If user selected [Remove] add /u (unregister) parameter (/unregister is not recognised sometimes)
        if ($Selection -eq 1)
        {
            $RegAsmArgs = @("/u") + $RegAsmArgs

			# Detect if finsql.exe is running (can't unregister until closed)
            $FinsqlProcess = Get-Process finsql -ErrorAction SilentlyContinue
            if($FinsqlProcess)
            {
                throw "Close all finsql.exe instances before removing OutlookCOMM"
            }
        }

        # Run regasm.exe with passed parameters
        & $RegAsmPath @RegAsmArgs
    }
    exit(0)
}