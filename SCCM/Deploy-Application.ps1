<#
.SYNOPSIS
	This script performs the installation or uninstallation of an application(s).
.DESCRIPTION
	The script is provided as a template to perform an install or uninstall of an application(s).
	The script either performs an "Install" deployment type or an "Uninstall" deployment type.
	The install deployment type is broken down into 3 main sections/phases: Pre-Install, Install, and Post-Install.
	The script dot-sources the AppDeployToolkitMain.ps1 script which contains the logic and functions required to install or uninstall an application.
.PARAMETER DeploymentType
	The type of deployment to perform. Default is: Install.
.PARAMETER DeployMode
	Specifies whether the installation should be run in Interactive, Silent, or NonInteractive mode. Default is: Interactive. Options: Interactive = Shows dialogs, Silent = No dialogs, NonInteractive = Very silent, i.e. no blocking apps. NonInteractive mode is automatically set if it is detected that the process is not user interactive.
.PARAMETER AllowRebootPassThru
	Allows the 3010 return code (requires restart) to be passed back to the parent process (e.g. SCCM) if detected from an installation. If 3010 is passed back to SCCM, a reboot prompt will be triggered.
.PARAMETER TerminalServerMode
	Changes to "user install mode" and back to "user execute mode" for installing/uninstalling applications for Remote Destkop Session Hosts/Citrix servers.
.PARAMETER DisableLogging
	Disables logging to file for the script. Default is: $false.
.EXAMPLE
    powershell.exe -Command "& { & '.\Deploy-Application.ps1' -DeployMode 'Silent'; Exit $LastExitCode }"
.EXAMPLE
    powershell.exe -Command "& { & '.\Deploy-Application.ps1' -AllowRebootPassThru; Exit $LastExitCode }"
.EXAMPLE
    powershell.exe -Command "& { & '.\Deploy-Application.ps1' -DeploymentType 'Uninstall'; Exit $LastExitCode }"
.EXAMPLE
    Deploy-Application.exe -DeploymentType "Install" -DeployMode "Silent"
.NOTES
	Toolkit Exit Code Ranges:
	60000 - 68999: Reserved for built-in exit codes in Deploy-Application.ps1, Deploy-Application.exe, and AppDeployToolkitMain.ps1
	69000 - 69999: Recommended for user customized exit codes in Deploy-Application.ps1
	70000 - 79999: Recommended for user customized exit codes in AppDeployToolkitExtensions.ps1
.LINK 
	http://psappdeploytoolkit.com
#>
[CmdletBinding()]
Param (
	[Parameter(Mandatory=$false)]
	[ValidateSet('Install','Uninstall')]
	[string]$DeploymentType = 'Install',
	[Parameter(Mandatory=$false)]
	[ValidateSet('Interactive','Silent','NonInteractive')]
	[string]$DeployMode = 'Interactive',
	[Parameter(Mandatory=$false)]
	[switch]$AllowRebootPassThru = $false,
	[Parameter(Mandatory=$false)]
	[switch]$TerminalServerMode = $false,
	[Parameter(Mandatory=$false)]
	[switch]$DisableLogging = $false
)

Try {
	## Set the script execution policy for this process
	Try { Set-ExecutionPolicy -ExecutionPolicy 'ByPass' -Scope 'Process' -Force -ErrorAction 'Stop' } Catch {}
	
	##*===============================================
	##* VARIABLE DECLARATION
	##*===============================================
	## Variables: Application
	[string]$appVendor = 'Microsoft'
	[string]$appName = 'Office 365 ProPlus'
	[string]$appVersion = '1708.8431.2153'
    # change for 32/64 bitness
    [string]$appArch = 'x86'
	[string]$appLang = 'EN'
	[string]$appRevision = '01'
	[string]$appScriptVersion = '2.0.3'
	[string]$appScriptDate = '01/17/2018'
	[string]$appScriptAuthor = 'Josh S'
    [string]$EnbridgeLogPath = $env:windir + '\ENBLogs'
    [string]$Office365MigrationLogPath = $EnbridgeLogPath + '\Office365Migration'
    [string]$Office365MigrationLogPathUninstall = $Office365MigrationLogPath + '\Uninstall'
	##*===============================================
	## Variables: Install Titles (Only set here to override defaults set by the toolkit)
	[string]$installName = ''
	[string]$installTitle = ''
	
	##* Do not modify section below
	#region DoNotModify
	
	## Variables: Exit Code
	[int32]$mainExitCode = 0
	
	## Variables: Script
	[string]$deployAppScriptFriendlyName = 'Deploy Application'
	[version]$deployAppScriptVersion = [version]'3.6.9'
	[string]$deployAppScriptDate = '02/06/2016'
	[hashtable]$deployAppScriptParameters = $psBoundParameters
	
	## Variables: Environment
	If (Test-Path -LiteralPath 'variable:HostInvocation') { $InvocationInfo = $HostInvocation } Else { $InvocationInfo = $MyInvocation }
	[string]$scriptDirectory = Split-Path -Path $InvocationInfo.MyCommand.Definition -Parent
	
	## Dot source the required App Deploy Toolkit Functions
	Try {
		[string]$moduleAppDeployToolkitMain = "$scriptDirectory\AppDeployToolkit\AppDeployToolkitMain.ps1"
		If (-not (Test-Path -LiteralPath $moduleAppDeployToolkitMain -PathType 'Leaf')) { Throw "Module does not exist at the specified location [$moduleAppDeployToolkitMain]." }
		If ($DisableLogging) { . $moduleAppDeployToolkitMain -DisableLogging } Else { . $moduleAppDeployToolkitMain }
	}
	Catch {
		If ($mainExitCode -eq 0){ [int32]$mainExitCode = 60008 }
		Write-Error -Message "Module [$moduleAppDeployToolkitMain] failed to load: `n$($_.Exception.Message)`n `n$($_.InvocationInfo.PositionMessage)" -ErrorAction 'Continue'
		## Exit the script, returning the exit code to SCCM
		If (Test-Path -LiteralPath 'variable:HostInvocation') { $script:ExitCode = $mainExitCode; Exit } Else { Exit $mainExitCode }
	}
	
    
	##*===============================================
	##* END VARIABLE DECLARATION
	##*===============================================
		
    If ((gwmi win32_operatingsystem | SELECT osarchitecture).osarchitecture -eq "64-bit") {
        $bitness = "64bit"
        Write-Log -Message '[BITNESS] detected a 64-bit Operating System.' -Source $deployAppScriptFriendlyName
    }
    Else {
        $bitness = "32bit"
        Write-Log -Message '[BITNESS] detected a 32-bit Operating System.' -Source $deployAppScriptFriendlyName
    }



	If ($deploymentType -ine 'Uninstall') {

		##*===============================================
		##* PRE-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Pre-Installation'
        Write-Log -Message 'Starting Pre-Installation Stage.' -Source $deployAppScriptFriendlyName

        # Set the initial Office folder
        [string] $dirOffice = Join-Path -Path "$envProgramFilesX86" -ChildPath "Microsoft Office"
        [string] $dirOfficeX64 = Join-Path -Path "$envProgramFiles" -ChildPath "Microsoft Office"
        [string] $dirOfficeC2R_2010 = Join-Path -Path "$envProgramFilesX86" -ChildPath "Microsoft Office 14"
        [string] $dirOfficeC2R_2010X64 = Join-Path -Path "$envProgramFiles" -ChildPath "Microsoft Office 14"
        [string] $dirOfficeC2R_2013 = Join-Path -Path "$envProgramFilesX86" -ChildPath "Microsoft Office 15"
        [string] $dirOfficeC2R_2013X64 = Join-Path -Path "$envProgramFiles" -ChildPath "Microsoft Office 15"      
        [string] $dirOfficeC2R_2016 = Join-Path -Path "$envProgramFilesX86" -ChildPath "Microsoft Office"
        [string] $dirOfficeC2R_2016X64 = Join-Path -Path "$envProgramFiles" -ChildPath "Microsoft Office"       
        
        Show-InstallationPrompt `
            -Message "`n`nReady to install Microsoft Office 365 ProPlus 2016! `n`nPlease note that the installation will take approximately 1 hour.`n`nAll previous versions of Microsoft Office will be uninstalled.`n`nAny conflicting applications will be closed.  `n`nPlease SAVE all of your current work and click OK to continue with the installation.`n`nA reboot will be required to finalize the installation.`n`n" `
            -ButtonRightText "OK" `
            -Icon Information `
            -Timeout 600 `
            -ExitOnTimeout $false
        
        Start-Sleep 5

        ## Show Welcome Message
        Show-InstallationWelcome -CloseApps 'ose,osppsvc,sppsvc,msoia,excel,groove,onenote,infopath,onenote,outlook,mspub,powerpnt,winword,winproj,visio,iexplore,msaccess,lync' -ForceCloseAppsCountdown 300 -CheckDiskSpace -PersistPrompt  
       
        ## Show Progress Message
        Show-InstallationProgress -StatusMessage "`n`nMicrosoft Office 365 Installation is in progress.`n`nThe Installation will take approximately 1 hour to complete.`n`nA reboot will be required to complete." -TopMost $False

        # Remove any previous version of Office (if required)
        [string[]]$officeExecutables = 'excel.exe', 'groove.exe', 'infopath.exe', 'onenote.exe', 'outlook.exe', 'mspub.exe', 'powerpnt.exe', 'winword.exe', 'lync.exe', 'msaccess.exe'

        #construct offscrub path
        $Offscrub03 = '"' + $dirSupportFiles + '\OffScrub03.vbs' + '"'
        $Offscrub07 = '"' + $dirSupportFiles + '\OffScrub07.vbs' + '"'
        $Offscrub10 = '"' + $dirSupportFiles + '\OffScrub10.vbs' + '"'
        $Offscrub15msi = '"' + $dirSupportFiles + '\OffScrub_O15msi.vbs' + '"'
        $Offscrub16msi = '"' + $dirSupportFiles + '\OffScrub_O16msi.vbs' + '"'
        $Offscrubc2r = '"' + $dirSupportFiles + '\OffScrubc2r.vbs' + '"'

        # create the main log folder if it doesn't already exist
        If(!(test-path $Office365MigrationLogPath)){
        New-Item -ItemType Directory -Force -Path $Office365MigrationLogPath
        }

        Write-Log -Message 'Starting To detect and scrub previous versions of Microsoft Office.' -Source $deployAppScriptFriendlyName

               
        ForEach ($officeExecutable in $officeExecutables) {
            If (Test-Path -Path (Join-Path -Path $dirOffice -ChildPath "Office11\$officeExecutable") -PathType Leaf) {
            Write-Log -Message 'Microsoft Office 2003 (32-bit) was detected. Will be uninstalled.' -Source $deployAppScriptFriendlyName
            ## Display Pre-Install cleanup Office 2003 
            Show-InstallationProgress -StatusMessage "Removing Microsoft Office 2003 (32-bit).  This may take some time to complete. Please wait…"
            Execute-Process -Path 'cscript.exe' -Parameters "$Offscrub03 ALL /S /Q /NoCancel /Bypass 1 /OSE /LOG $Office365MigrationLogPath\Office2003Scrub_32bit" -WindowStyle Hidden -IgnoreExitCodes '1,2,3,16,42'
            Write-Log -Message 'Finished scrubbing Microsoft Office 2003.' -Source $deployAppScriptFriendlyName
            Break
            }
        }
        
        ForEach ($officeExecutable in $officeExecutables) {
            If (Test-Path -Path (Join-Path -Path $dirOffice -ChildPath "Office12\$officeExecutable") -PathType Leaf) {
            Write-Log -Message 'Microsoft Office 2007 (32-bit) was detected. Will be uninstalled.' -Source $deployAppScriptFriendlyName
            ## Display Pre-Install cleanup Office 2007
            Show-InstallationProgress -StatusMessage "Removing Microsoft Office 2007 (32-bit).  This may take some time to complete. Please wait…"
            Execute-Process -Path 'cscript.exe' -Parameters "$Offscrub07 ALL /S /Q /NoCancel /Bypass 1 /OSE /LOG $Office365MigrationLogPath\Office2007Scrub_32bit" -WindowStyle Hidden -IgnoreExitCodes '1,2,3,16,42'
            Write-Log -Message 'Finished scrubbing Microsoft Office 2007.' -Source $deployAppScriptFriendlyName
            Break
            }
        }

        ## Uninstall Enterprise Vault for Office 2010
        Write-Log -Message 'Starting to Uninstall Enterprise Vault for Outlook 2010.  It may not be installed, but just to be sure.' -Source $deployAppScriptFriendlyName
        [string]$EV_MSI_PC = '{B1742DBF-903F-4416-9F32-8BB9F7B60B83}'
        Execute-Process -Path "$env:windir\system32\msiexec.exe" -Parameters "/x $EV_MSI_PC /qn REBOOT=ReallySuppress /l*v $Office365MigrationLogPath\SCCM_EnterpriseVaultForOutlook2010_1002OLD_Uninstall.log"
        Write-Log -Message 'Finished Uninstalling Enterprise Vault for Outlook 2010' -Source $deployAppScriptFriendlyName


        ForEach ($officeExecutable in $officeExecutables) {
            If (Test-Path -Path (Join-Path -Path $dirOffice -ChildPath "Office14\$officeExecutable") -PathType Leaf) {
            Write-Log -Message 'Microsoft Office 2010 (32-bit) was detected. Will be uninstalled.' -Source $deployAppScriptFriendlyName
            ## Display Pre-Install cleanup Office 2010
            Show-InstallationProgress -StatusMessage "Removing Microsoft Office 2010 (32-bit).  This may take some time to complete. Please wait…"
            Execute-Process -Path "cscript.exe" -Parameters "$Offscrub10 ALL /S /Q /NoCancel /Bypass 1 /OSE /LOG $Office365MigrationLogPath\Office2010Scrub_32bit" -WindowStyle Hidden -IgnoreExitCodes '1,2,3,16,42'
            Write-Log -Message 'Finished scrubbing Microsoft Office 2010.' -Source $deployAppScriptFriendlyName
            Break
            }
        }
        
        If ($bitness -eq "64bit") {
            ForEach ($officeExecutable in $officeExecutables) {
                If (Test-Path -Path (Join-Path -Path $dirOfficeX64 -ChildPath "Office14\$officeExecutable") -PathType Leaf) {
                Write-Log -Message 'Microsoft Office 2010 (64-bit) was detected. Will be uninstalled.' -Source $deployAppScriptFriendlyName
                ## Display Pre-Install cleanup Office 2010
                Show-InstallationProgress -StatusMessage "Removing Microsoft Office 2010 (64-bit).  This may take some time to complete. Please wait…"
                Execute-Process -Path "cscript.exe" -Parameters "$Offscrub10 ALL /S /Q /NoCancel /Bypass 1 /OSE /LOG $Office365MigrationLogPath\Office2010Scrub_64bit" -WindowStyle Hidden -IgnoreExitCodes '1,2,3,16,42'
                Write-Log -Message 'Finished scrubbing Microsoft Office 2010.' -Source $deployAppScriptFriendlyName
                Break
                }
            }
        }


        ## Click-to-Run Office 2010 [32-bit]
        # ForEach ($officeExecutable in $officeExecutables) {
        #    If (Test-Path -Path (Join-Path -Path $dirOfficeC2R_2010 -ChildPath "root\Office14\$officeExecutable") -PathType Leaf) {
        #    Write-Log -Message 'Microsoft Office 2010 C2R (32-bit) was detected. Will be uninstalled.' -Source $deployAppScriptFriendlyName
        #    ## Display Pre-Install cleanup Office 2010 C2R
        #    Show-InstallationProgress -StatusMessage "Removing Microsoft Office 2010 C2R (32-bit).  This may take some time to complete. Please wait…"
        #    Execute-Process -Path "cscript.exe" -Parameters "$Offscrubc2r ALL /S /Q /NoCancel /Bypass 1 /OSE /LOG $Office365MigrationLogPath\Office2010C2RScrub_32bit" -WindowStyle Hidden -IgnoreExitCodes '1,2,3,8,42,34,67'
        #    Write-Log -Message 'Finished scrubbing Microsoft Office 2016 C2R (32-bit).' -Source $deployAppScriptFriendlyName
        #    Break
        #    }
        #}

        #If ($bitness -eq "64bit") {
        ## Click-to-Run Office 2010 [64-bit]
        #    If ((gwmi win32_operatingsystem | SELECT osarchitecture).osarchitecture -eq "64-bit") {
        #        ForEach ($officeExecutable in $officeExecutables) {
        #            If (Test-Path -Path (Join-Path -Path $dirOfficeC2R_2010X64 -ChildPath "root\Office14\$officeExecutable") -PathType Leaf) {
        #            Write-Log -Message 'Microsoft Office 2010 C2R (64-bit) was detected. Will be uninstalled.' -Source $deployAppScriptFriendlyName            
        #            ## Display Pre-Install cleanup Office 2010
        #            Show-InstallationProgress -StatusMessage "Removing Microsoft Office 2010 C2R (64-bit).  This may take some time to complete. Please wait…"
        #            Execute-Process -Path "cscript.exe" -Parameters "$Offscrubc2r ALL /S /Q /NoCancel /Bypass 1 /OSE /LOG $Office365MigrationLogPath\Office2010C2RScrub_64bit" -WindowStyle Hidden -IgnoreExitCodes '1,2,3,8,42,34,67'
        #            Write-Log -Message 'Finished scrubbing Microsoft Office 2010 C2R (64-bit).' -Source $deployAppScriptFriendlyName
        #            Break
        #            }
        #        }
        #    }	
        #}

        # OneDrive for Business 2013 [32-bit]
        # Click-to-RunL Skype for Business 2015 (aka Lync) [32-bit]
        ForEach ($officeExecutable in $officeExecutables) {
            If (Test-Path -Path (Join-Path -Path $dirOffice -ChildPath "Office15\$officeExecutable") -PathType Leaf) {
            Write-Log -Message 'Removing Microsoft Office 2013 (32-bit) was detected. Will be uninstalled.' -Source $deployAppScriptFriendlyName
            ## Display Pre-Install cleanup Office 2013 - which is likely just OneDrive
            Show-InstallationProgress -StatusMessage "Removing Microsoft Office 2013 (32-bit) - OneDrive.  This may take some time to complete. Please wait…"
            Execute-Process -Path "cscript.exe" -Parameters "$Offscrub15msi ALL /S /Q /NoCancel /Bypass 1 /OSE /LOG $Office365MigrationLogPath\Office2013Scrub_32bit" -WindowStyle Hidden -IgnoreExitCodes '1,2,3,16,42'
            Write-Log -Message 'Finished scrubbing Microsoft Office 2013.' -Source $deployAppScriptFriendlyName
            Break
            }
        }

        # OneDrive for Business 2013 [64-bit]
        # Click-to-RunL Skype for Business 2015 (aka Lync) [64-bit]
        If ($bitness -eq "64bit") {
            ForEach ($officeExecutable in $officeExecutables) {
                If (Test-Path -Path (Join-Path -Path $dirOfficeX64 -ChildPath "Office15\$officeExecutable") -PathType Leaf) {
                Write-Log -Message 'Removing Microsoft Office 2013 (64-bit) was detected. Will be uninstalled.' -Source $deployAppScriptFriendlyName
                ## Display Pre-Install cleanup Office 2013 - which is likely just OneDrive
                Show-InstallationProgress -StatusMessage "Removing Microsoft Office 2013 (64-bit) - OneDrive.  This may take some time to complete. Please wait…"
                Execute-Process -Path "cscript.exe" -Parameters "$Offscrub15msi ALL /S /Q /NoCancel /Bypass 1 /OSE /LOG $Office365MigrationLogPath\Office2013Scrub_64bit" -WindowStyle Hidden -IgnoreExitCodes '1,2,3,16,42'
                Write-Log -Message 'Finished scrubbing Microsoft Office 2013.' -Source $deployAppScriptFriendlyName
                Break
                }
            }
        }


        ForEach ($officeExecutable in $officeExecutables) {
            If (Test-Path -Path (Join-Path -Path $dirOfficeC2R_2013 -ChildPath "root\Office15\$officeExecutable") -PathType Leaf) {
            Write-Log -Message 'Microsoft Office 2013 C2R (32-bit) was detected. Will be uninstalled.' -Source $deployAppScriptFriendlyName
            ## Display Pre-Install cleanup Office 2013 C2R - which is likely just Skype
            Show-InstallationProgress -StatusMessage "Removing Microsoft Office 2013 C2R (32-bit) - Skype.  This may take some time to complete. Please wait…"
            Execute-Process -Path "cscript.exe" -Parameters "$Offscrubc2r ALL /S /Q /NoCancel /Bypass 1 /OSE /LOG $Office365MigrationLogPath\Office2013C2RScrub_32bit" -WindowStyle Hidden -IgnoreExitCodes '1,2,3,8,42,34,67'
            Write-Log -Message 'Finished scrubbing Microsoft Office 2013 C2R (32-bit).' -Source $deployAppScriptFriendlyName
            Break
            }
        }

        If ($bitness -eq "64bit") {
            ForEach ($officeExecutable in $officeExecutables) {
                If (Test-Path -Path (Join-Path -Path $dirOfficeC2R_2013X64 -ChildPath "root\Office15\$officeExecutable") -PathType Leaf) {
                Write-Log -Message 'Microsoft Office 2013 C2R (64-bit) was detected. Will be uninstalled.' -Source $deployAppScriptFriendlyName
                ## Display Pre-Install cleanup Office 2013 C2R - which is likely just Skype
                Show-InstallationProgress -StatusMessage "Removing Microsoft Office 2013 C2R (64-bit) - Skype.  This may take some time to complete. Please wait…"
                Execute-Process -Path "cscript.exe" -Parameters "$Offscrubc2r ALL /S /Q /NoCancel /Bypass 1 /OSE /LOG $Office365MigrationLogPath\Office2013C2RScrub_64bit" -WindowStyle Hidden -IgnoreExitCodes '1,2,3,8,42,34,67'
                Write-Log -Message 'Finished scrubbing Microsoft Office 2013 C2R (64-bit).' -Source $deployAppScriptFriendlyName
                Break
                }
            }		
		}

        # force delete of start manu folders/shortcuts left behind
        If (test-path "c:\programdata\microsoft\windows\start menu\programs\microsoft office 2013\"){
            remove-item "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2013\" -recurse
        }

        # Office 2016 [32-bit]
        ForEach ($officeExecutable in $officeExecutables) {
            If (Test-Path -Path (Join-Path -Path $dirOffice -ChildPath "Office16\$officeExecutable") -PathType Leaf) {
            Write-Log -Message 'Removing Microsoft Office 2016 (32-bit) was detected. Will be uninstalled.' -Source $deployAppScriptFriendlyName
            ## Display Pre-Install cleanup Office 2016
            Show-InstallationProgress -StatusMessage "Removing Microsoft Office 2016 (32-bit).  This may take some time to complete. Please wait…"
            Execute-Process -Path "cscript.exe" -Parameters "$Offscrub16msi ALL /S /Q /NoCancel /Bypass 1 /OSE /LOG $Office365MigrationLogPath\Office2016Scrub_32bit" -WindowStyle Hidden -IgnoreExitCodes '1,2,3,16,42'
            Write-Log -Message 'Finished scrubbing Microsoft Office 2016 (32-bit).' -Source $deployAppScriptFriendlyName
            Break
            }
        }

        # Office 2016 [64-bit]
        If ($bitness -eq "64bit") {
            ForEach ($officeExecutable in $officeExecutables) {
                If (Test-Path -Path (Join-Path -Path $dirOfficeX64 -ChildPath "Office16\$officeExecutable") -PathType Leaf) {
                Write-Log -Message 'Removing Microsoft Office 2016 (64-bit) was detected. Will be uninstalled.' -Source $deployAppScriptFriendlyName
                ## Display Pre-Install cleanup Office 2016
                Show-InstallationProgress -StatusMessage "Removing Microsoft Office 2016 (64-bit).  This may take some time to complete. Please wait…"
                Execute-Process -Path "cscript.exe" -Parameters "$Offscrub16msi ALL /S /Q /NoCancel /Bypass 1 /OSE /LOG $Office365MigrationLogPath\Office2016Scrub_64bit" -WindowStyle Hidden -IgnoreExitCodes '1,2,3,16,42'
                Write-Log -Message 'Finished scrubbing Microsoft Office 2016 (64-bit).' -Source $deployAppScriptFriendlyName
                Break
                }
            }
        }

        # Click-to-Run Office 2016 [32-bit]
        ForEach ($officeExecutable in $officeExecutables) {
            If (Test-Path -Path (Join-Path -Path $dirOfficeC2R_2016 -ChildPath "root\Office16\$officeExecutable") -PathType Leaf) {
            Write-Log -Message 'Microsoft Office 2016 C2R (32-bit) was detected. Will be uninstalled.' -Source $deployAppScriptFriendlyName
            ## Display Pre-Install cleanup Office 2016 C2R
            Show-InstallationProgress -StatusMessage "Removing Microsoft Office 2016 C2R (32-bit).  This may take some time to complete. Please wait…"
            Execute-Process -Path "cscript.exe" -Parameters "$Offscrubc2r ALL /S /Q /NoCancel /Bypass 1 /OSE /LOG $Office365MigrationLogPath\Office2016C2RScrub_32bit" -WindowStyle Hidden -IgnoreExitCodes '1,2,3,8,42,34,67'
            Write-Log -Message 'Finished scrubbing Microsoft Office 2016 C2R (32-bit).' -Source $deployAppScriptFriendlyName
            Break
            }
        }

        # Click-to-Run Office 2016 [64-bit]
        If ($bitness -eq "64bit") {
            ForEach ($officeExecutable in $officeExecutables) {
                If (Test-Path -Path (Join-Path -Path $dirOfficeC2R_2016X64 -ChildPath "root\Office16\$officeExecutable") -PathType Leaf) {
                Write-Log -Message 'Microsoft Office 2016 C2R (64-bit) was detected. Will be uninstalled.' -Source $deployAppScriptFriendlyName            
                ## Display Pre-Install cleanup Office 2016
                Show-InstallationProgress -StatusMessage "Removing Microsoft Office 2016 C2R (64-bit).  This may take some time to complete. Please wait…"
                Execute-Process -Path "cscript.exe" -Parameters "$Offscrubc2r ALL /S /Q /NoCancel /Bypass 1 /OSE /LOG $Office365MigrationLogPath\Office2016C2RScrub_64bit" -WindowStyle Hidden -IgnoreExitCodes '1,2,3,8,42,34,67'
                Write-Log -Message 'Finished scrubbing Microsoft Office 2016 C2R (64-bit).' -Source $deployAppScriptFriendlyName
                Break
                }
            }
        }	


        Write-Log -Message 'Finished detecting and scrubbin previous versions of Microsoft Office.' -Source $deployAppScriptFriendlyName
        Write-Log -Message 'Finished Post-Installation Stage.' -Source $deployAppScriptFriendlyName

 
      	## Call the Exit-Script function to perform final cleanup operations
    	#Exit-Script -ExitCode $mainExitCode

		##*===============================================
		##* INSTALLATION 
		##*===============================================
		[string]$installPhase = 'Installation'
        Write-Log -Message 'Installation Installation Stage.' -Source $deployAppScriptFriendlyName

		Show-InstallationProgress -StatusMessage 'Installing Office 365 ProPlus 2016 (32-bit).  This may take some time to complete. Please wait…' -TopMost $True
        # Install Microsoft Office 365 32-bit
        Write-Log -Message 'Starting to Install Microsoft Office 365 Base Install - 32-bit Version.' -Source $deployAppScriptFriendlyName
        Execute-Process -Path "$dirFiles\Office365_32bit\Setup.exe" -Parameters "/configure InsOfficeAll.xml"
        Write-Log -Message 'Finished Installing Microsoft Office 365 Base Install - 32-bit Version.' -Source $deployAppScriptFriendlyName
                
        Write-Log -Message 'Finished Installation Stage.' -Source $deployAppScriptFriendlyName
                			
		##*===============================================
		##* POST-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Installation'
        Write-Log -Message 'Starting Post-Installation Stage.' -Source $deployAppScriptFriendlyName

        ## Let's make sure appropriate Office Registry keys that indicate bitness are in place.  Sometimes they aren't set.
        Write-Log -Message 'Starting to set registry keys to indicate bitness.' -Source $deployAppScriptFriendlyName
        If ($bitness -eq "64bit") {
            Write-Log -Message 'Since we have installed Office 32-bit on a 64-bit Windows OS, we need to set key [SOFTWARE\Wow6432Node\Microsoft\Office\16.0\Outlook\Bitness] to be equal to [x86].' -Source $deployAppScriptFriendlyName
            Set-RegistryKey -Key 'HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Office\16.0\Outlook' -Name 'Bitness' -Value x86 -Type 'String'
            Write-Log -Message 'Attempted to set registry key.' -Source $deployAppScriptFriendlyName
        }
        Else {
            Write-Log -Message 'Since we have installed Office 32-bit on a 32-bit Windows OS, we need to set key [SOFTWARE\Microsoft\Office\16.0\Outlook\Bitness] to be equal to [x86].' -Source $deployAppScriptFriendlyName
            Set-RegistryKey -Key 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Outlook' -Name 'Bitness' -Value x86 -Type 'String'
            Write-Log -Message 'Attempted to set registry key.' -Source $deployAppScriptFriendlyName
        }
        Write-Log -Message 'Finished setting registry keys to indicate bitness.' -Source $deployAppScriptFriendlyName

        ## Install Outlook Add-Ins
        Write-Log -Message 'Starting to Install Outlook Add-Ins.  Note that bitness of add-ins will need to match bitness of Office 365 (not OS).' -Source $deployAppScriptFriendlyName
        
        ## Display Add-in installations 
        Show-InstallationProgress -StatusMessage "Installing Office Add-Ins.  This may take some time to complete. Please wait…"

        ## Install EMM (OpenText Email Management for MSX v16.0.0
        ## Comment this part out for Legacy Spectra - They don't have EMM
        Write-Log -Message 'Starting to Install OpenText Email Management for MSX v16.0.0 - 32-bit Version.' -Source $deployAppScriptFriendlyName
        Execute-Process -Path "$env:windir\system32\msiexec.exe" -Parameters "/i $dirFiles\EMM_32bit\EmailManagementAddin_32_16.0.0.msi /qn REBOOT=ReallySuppress /l*v $Office365MigrationLogPath\SCCM_EmailManagement16x86_Install.log"
        Write-Log -Message 'Finished Installing OpenText Email Management for MSX v16.0.0 - 32-bit Version.' -Source $deployAppScriptFriendlyName

        ## Install Phishme (PhishMe Reporter v3.1.3.0 (Outlook Add-In)
        Write-Log -Message 'Starting to Install PhishMe Reporter v3.1.3.0' -Source $deployAppScriptFriendlyName
        Execute-Process -Path "$dirFiles\PhishMe\PhishME3130.exe" -Parameters "/S"
        Write-Log -Message 'Finished Installing PhishMe Reporter v3.1.3.0' -Source $deployAppScriptFriendlyName

        # Done Installing Outlook Add-Ins
        Write-Log -Message 'Finished Installing Outlook Add-Ins.' -Source $deployAppScriptFriendlyName

        # Done POST-INSTALLATION Stage
        Write-Log -Message 'Finished Post-Installation Stage.' -Source $deployAppScriptFriendlyName

        # copy marker file
        ## This script installs Office 365 [32-bit] so lets be sure to get 32-bit program files folder
        If ($bitness -eq "64bit") {$ProgFilePath=[Environment]::GetEnvironmentVariable("ProgramFiles(x86)")}
        Else {$ProgFilePath=[Environment]::GetEnvironmentVariable("ProgramFiles")}

        $fileSource=Split-Path -Path $MyInvocation.MyCommand.Path
        If (Test-Path "$ProgFilePath\Microsoft Office") {Copy-Item "$fileSource\EnbridgeOffice365MigrationE1.txt" "$ProgFilePath\Microsoft Office"}
        If (Test-Path "$ProgFilePath\Microsoft Office\EnbridgeOffice365MigrationE1.txt") {Write-Log -Message 'Copied marker file successfully. This is important because it is used for SCCM Detection method.' -Source $deployAppScriptFriendlyName}
        Else {Write-Log -Message 'Problem writing marker file.  This is important because it is used for SCCM Detection method.' -Source $deployAppScriptFriendlyName}

		## Display a message at the end of the install

        ### put in a reboot here.  5 minutes. no deferrals, just and OK button
        # If (-not $useDefaultMsi) { Show-InstallationPrompt -Message "`n`nMicrosoft Office 365 ProPlus 2016 Installation is now complete!`n`nA restart is required to finalize the installation." -ButtonRightText 'OK' -Icon Exclamation -Timeout 300 }
        Write-Log -Message 'Display InstallationRestartPrompt to user notifying them install has finished, and reboot required to complete.' -Source $deployAppScriptFriendlyName
        If (-not $useDefaultMsi) { Show-InstallationRestartPrompt -CountdownSeconds 300 -CountdownNoHideSeconds 300}
  
	}
	ElseIf ($deploymentType -ieq 'Uninstall')
	{
		##*===============================================
		##* PRE-UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Pre-Uninstallation'
        Write-Log -Message 'Starting Pre-Uninstallation Stage.' -Source $deployAppScriptFriendlyName

        # create the main log folder for uninstall if it doesn't already exist
        If(!(test-path $Office365MigrationLogPathUninstall)){
        New-Item -ItemType Directory -Force -Path $Office365MigrationLogPathUninstall
        }

        # Set the initial Office folder
        [string] $dirOffice = Join-Path -Path "$envProgramFilesX86" -ChildPath "Microsoft Office"
        [string] $dirOfficeX64 = Join-Path -Path "$envProgramFiles" -ChildPath "Microsoft Office"
        [string] $dirOfficeC2R_2016 = Join-Path -Path "$envProgramFilesX86" -ChildPath "Microsoft Office"
        [string] $dirOfficeC2R_2016X64 = Join-Path -Path "$envProgramFiles" -ChildPath "Microsoft Office"     


        Show-InstallationPrompt `
            -Message "`n`nReady to Uninstall Microsoft Office 365 ProPlus 2016! `n`nPlease note that the removal will take approximately 15minutes.`n`nAny conflicting applications will be closed.  `n`nPlease SAVE all of your current work and click OK to continue with the removal.`n`nA reboot will be required when finished.`n`n" `
            -ButtonRightText "OK" `
            -Icon Information `
            -Timeout 300 `
            -ExitOnTimeout $false

        ## Show Welcome Message, 
        Show-InstallationWelcome -CloseApps 'ose,osppsvc,sppsvc,msoia,excel,groove,onenote,infopath,onenote,outlook,mspub,powerpnt,winword,winproj,visio,iexplore,msaccess,lync' -ForceCloseAppsCountdown 300 -PersistPrompt 
        
        # Show-InstallationProgress
        Show-InstallationProgress -StatusMessage "`n`nMicrosoft Office 365 Removal is in progress.`n`nThis will take approximately 15 minutes to complete.`n`nA reboot will be required" -TopMost $False

        Write-Log -Message 'Finished Pre-Uninstallation Stage.' -Source $deployAppScriptFriendlyName
		
		##*===============================================
		##* UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Uninstallation'
        Write-Log -Message 'Starting Uninstallation Stage.' -Source $deployAppScriptFriendlyName
		
		## Handle Zero-Config MSI Uninstallations
		If ($useDefaultMsi) {
			[hashtable]$ExecuteDefaultMSISplat =  @{ Action = 'Uninstall'; Path = $defaultMsiFile }; If ($defaultMstFile) { $ExecuteDefaultMSISplat.Add('Transform', $defaultMstFile) }
			Execute-MSI @ExecuteDefaultMSISplat
		}
		
		# <Perform Uninstallation tasks here>
        Show-InstallationProgress -StatusMessage 'Uninstalling Microsoft Office 365 ProPlus Add-ins and Base Components. This may take some time. Please wait…'

        ## Uninstall Outlook Add-Ins
        Write-Log -Message 'Starting to Uninstall Outlook Add-Ins.' -Source $deployAppScriptFriendlyName

        ## Uninstall EMM (OpenText Email Management for MSX v16.0.0
        ## Comment this out for Legacy Spectra.  They don't use EMM
        Write-Log -Message 'Starting to Uninstall OpenText Email Management for MSX v16.0.0 - 32-bit Version.' -Source $deployAppScriptFriendlyName
        [string]$EMM_MSI_PC = '{E9B9634E-6714-4615-B9EE-08410A441C99}'
        Execute-Process -Path "$env:windir\system32\msiexec.exe" -Parameters "/x $EMM_MSI_PC /qn REBOOT=ReallySuppress /l*v $Office365MigrationLogPathUninstall\SCCM_EmailManagement16x86_Uninstall.log"
        Write-Log -Message 'Finished Uninstalling OpenText Email Management for MSX v16.0.0 - 32-bit Version.' -Source $deployAppScriptFriendlyName

        ## Uninstall Phishme (PhishMe Reporter v3.1.3.0 (Outlook Add-In)
        Write-Log -Message 'Starting to Uninstall PhishMe Reporter v3.1.3.0' -Source $deployAppScriptFriendlyName
        [string]$PhishMe_MSI_PC = '{6D7D704D-7278-4AF2-AA32-5C285268409D}'
        Execute-Process -Path "$env:windir\system32\msiexec.exe" -Parameters "/x $PhishMe_MSI_PC /qn REBOOT=ReallySuppress /l*v $Office365MigrationLogPathUninstall\SCCM_PhishMe3130_Uninstall.log"
        Write-Log -Message 'Finished Uninstalling PhishMe Reporter v3.1.3.0' -Source $deployAppScriptFriendlyName

        # Done Uninstalling Outlook Add-Ins
        Write-Log -Message 'Finished Installing Outlook Add-Ins.' -Source $deployAppScriptFriendlyName

        # Uninstall Microsoft Office 365 32-bit Base Components
        Write-Log -Message 'Starting to Uninstall Microsoft Office 365 Base Install - 32-bit Version.' -Source $deployAppScriptFriendlyName
        Execute-Process -Path "$dirFiles\Office365_32bit\Setup.exe" -Parameters "/configure RemoveOffice.xml"
   	    Write-Log -Message 'Finished Uninstalling Microsoft Office 365 Base Install - 32-bit Version.' -Source $deployAppScriptFriendlyName
		
        Write-Log -Message 'Finished Uninstallation Stage.' -Source $deployAppScriptFriendlyName

		##*===============================================
		##* POST-UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Uninstallation'
		Write-Log -Message 'Starting Post-Uninstallation Stage.' -Source $deployAppScriptFriendlyName

        ### DOUBLE SCRUB!!  Probably not needed.  comment out
        
        ## Office 2016 [32-bit]
        #ForEach ($officeExecutable in $officeExecutables) {
        #    If (Test-Path -Path (Join-Path -Path $dirOffice -ChildPath "Office16\$officeExecutable") -PathType Leaf) {
        #    Write-Log -Message 'Removing Microsoft Office 2016 (32-bit) was detected. Will be uninstalled.' -Source $deployAppScriptFriendlyName
        #    ## Display Pre-Install cleanup Office 2016
        #    Show-InstallationProgress -StatusMessage "Removing Microsoft Office 2016 (32-bit).  This may take some time to complete. Please wait…"
        #    Execute-Process -Path "cscript.exe" -Parameters "$Offscrub16msi ALL /S /Q /NoCancel /Bypass 1 /OSE /LOG $Office365MigrationLogPathUninstall\Office2016Scrub_32bit" -WindowStyle Hidden -IgnoreExitCodes '1,2,3,16,42'
        #    Write-Log -Message 'Finished scrubbing Microsoft Office 2016 (32-bit).' -Source $deployAppScriptFriendlyName
        #    Break
        #    }
        #}

        ## Office 2016 [64-bit]
        #If ($bitness -eq "64bit") {
        #    ForEach ($officeExecutable in $officeExecutables) {
        #        If (Test-Path -Path (Join-Path -Path $dirOfficeX64 -ChildPath "Office16\$officeExecutable") -PathType Leaf) {
        #        Write-Log -Message 'Removing Microsoft Office 2016 (64-bit) was detected. Will be uninstalled.' -Source $deployAppScriptFriendlyName
        #        ## Display Pre-Install cleanup Office 2016
        #        Show-InstallationProgress -StatusMessage "Removing Microsoft Office 2016 (64-bit).  This may take some time to complete. Please wait…"
        #        Execute-Process -Path "cscript.exe" -Parameters "$Offscrub16msi ALL /S /Q /NoCancel /Bypass 1 /OSE /LOG $Office365MigrationLogPathUninstall\Office2016Scrub_64bit" -WindowStyle Hidden -IgnoreExitCodes '1,2,3,16,42'
        #        Write-Log -Message 'Finished scrubbing Microsoft Office 2016 (64-bit).' -Source $deployAppScriptFriendlyName
        #        Break
        #        }
        #    }
        #}

        ## Click-to-Run Office 2016 [32-bit]
        #ForEach ($officeExecutable in $officeExecutables) {
        #    If (Test-Path -Path (Join-Path -Path $dirOfficeC2R_2016 -ChildPath "root\Office16\$officeExecutable") -PathType Leaf) {
        #    Write-Log -Message 'Microsoft Office 2016 C2R (32-bit) was detected. Will be uninstalled.' -Source $deployAppScriptFriendlyName
        #    ## Display Pre-Install cleanup Office 2016 C2R
        #    Show-InstallationProgress -StatusMessage "Removing Microsoft Office 2016 C2R (32-bit).  This may take some time to complete. Please wait…"
        #    Execute-Process -Path "cscript.exe" -Parameters "$Offscrubc2r ALL /S /Q /NoCancel /Bypass 1 /OSE /LOG $Office365MigrationLogPathUninstall\Office2016C2RScrub_32bit" -WindowStyle Hidden -IgnoreExitCodes '1,2,3,8,42,34,67'
        #    Write-Log -Message 'Finished scrubbing Microsoft Office 2016 C2R (32-bit).' -Source $deployAppScriptFriendlyName
        #    Break
        #    }
        #}

        ## Click-to-Run Office 2016 [64-bit]
        #If ($bitness -eq "64bit") {
        #    ForEach ($officeExecutable in $officeExecutables) {
        #        If (Test-Path -Path (Join-Path -Path $dirOfficeC2R_2016X64 -ChildPath "root\Office16\$officeExecutable") -PathType Leaf) {
        #        Write-Log -Message 'Microsoft Office 2016 C2R (64-bit) was detected. Will be uninstalled.' -Source $deployAppScriptFriendlyName            
        #        ## Display Pre-Install cleanup Office 2016
        #        Show-InstallationProgress -StatusMessage "Removing Microsoft Office 2016 C2R (64-bit).  This may take some time to complete. Please wait…"
        #        Execute-Process -Path "cscript.exe" -Parameters "$Offscrubc2r ALL /S /Q /NoCancel /Bypass 1 /OSE /LOG $Office365MigrationLogPathUninstall\Office2016C2RScrub_64bit" -WindowStyle Hidden -IgnoreExitCodes '1,2,3,8,42,34,67'
        #        Write-Log -Message 'Finished scrubbing Microsoft Office 2016 C2R (64-bit).' -Source $deployAppScriptFriendlyName
        #        Break
        #        }
        #    }
        #}	

		## <Perform Post-Uninstallation tasks here>
        
        ## This script uninstalls Office 365 [32-bit] so lets be sure to get 32-bit program files folder
        If ($bitness -eq "64bit") {$ProgFilePath=[Environment]::GetEnvironmentVariable("ProgramFiles(x86)")}
        Else {$ProgFilePath=[Environment]::GetEnvironmentVariable("ProgramFiles")}

        If (Test-Path "$ProgFilePath\Microsoft Office\EnbridgeOffice365MigrationE1.txt") {
            Remove-Item "$ProgFilePath\Microsoft Office\EnbridgeOffice365MigrationE1.txt"
            Write-Log -Message "Found and Deleted marker file successfully. This is important because it is used for SCCM Detection method." -Source $deployAppScriptFriendlyName
        }
        Else {
            Write-Log -Message "Couldn't find marker file used for SCCM detection method.  Fine.  We were going to delete it anyways." -Source $deployAppScriptFriendlyName
        }


        ## Let's make sure appropriate Office Registry keys that indicate bitness are removed.  Sometimes they aren't removed properly. 
        Write-Log -Message 'Starting to remove registry keys to that indicate bitness.' -Source $deployAppScriptFriendlyName
        If ($bitness -eq "64bit") {
            Write-Log -Message 'Since we have uninstalled Office 32-bit on a 64-bit Windows OS, we need to remove key [SOFTWARE\Wow6432Node\Microsoft\Office\16.0\Outlook\Bitness].' -Source $deployAppScriptFriendlyName
            Remove-RegistryKey -Key 'HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Office\16.0\Outlook' -Name 'Bitness'
            Write-Log -Message 'Attempted to remove registry key.' -Source $deployAppScriptFriendlyName
        }
        Else {
            Write-Log -Message 'Since we have uninstalled Office 32-bit on a 32-bit Windows OS, we need to remove key [SOFTWARE\Microsoft\Office\16.0\Outlook\Bitness].' -Source $deployAppScriptFriendlyName
            Remove-RegistryKey -Key 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Outlook' -Name 'Bitness'
            Write-Log -Message 'Attempted to remove registry key.' -Source $deployAppScriptFriendlyName
        }
        Write-Log -Message 'Finished removing registry keys to that indicate bitness.' -Source $deployAppScriptFriendlyName
        
        # prompt the users
        Write-Log -Message 'Display InstallationRestartPrompt to user notifying them uninstall has finished, and reboot required to complete.' -Source $deployAppScriptFriendlyName
        Show-InstallationPrompt -Message "`n`nMicrosoft Office 365 ProPlus Uninstall has been completed.`n`nYour Computer will need to be restarted, at your convenience, in order to apply all changes.`n`nPress OK button to exit" `
            -ButtonRightText "OK" `
            -Icon Information `
            -Timeout 120 `
            -ExitOnTimeout $false

        Write-Log -Message 'Finished Post-Uninstallation Stage.' -Source $deployAppScriptFriendlyName
		
	}
	
	##*===============================================
	##* END SCRIPT BODY
	##*===============================================
	
	## Call the Exit-Script function to perform final cleanup operations
	Exit-Script -ExitCode $mainExitCode
}
Catch {
	[int32]$mainExitCode = 60001
	[string]$mainErrorMessage = "$(Resolve-Error)"
	Write-Log -Message $mainErrorMessage -Severity 3 -Source $deployAppScriptFriendlyName
	Show-DialogBox -Text $mainErrorMessage -Icon 'Stop'
	Exit-Script -ExitCode $mainExitCode
}
