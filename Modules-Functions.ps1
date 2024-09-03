Function Install-RequiredModules {
    Param (
        [array] $requiredModules
    )

    ForEach ($module in $requiredModules) {
        $installedModules = Get-Module -ListAvailable -Name $module | Sort-Object Version -Descending
        $latestModule = Find-Module -Name $module

        # Check if module is installed
        If (!$installedModules) {
            Try {
                Install-Module -Name $module -Force -Scope CurrentUser -ErrorAction Stop
                Write-Output "Installed $module successfully."
            }
            Catch {
                If ($logFile) {
                    Add-Content -Path $logFile -Value "ERROR: Failed to install $module. $_"
                }
                If (Show-Error) {
                    Show-Error -errorMsg "Error: Failed to install module. $_"
                }
                Write-Warning "Failed to install $module. Error: $_"
                Exit
            }
        }
        ElseIf ($installedModules[0].Version -lt $latestModule.Version) {
            Try {
                Install-Module -Name $module -Force -Scope CurrentUser -ErrorAction Stop
                Write-Output "Updated $module to version $($latestModule.Version) successfully."

                # Remove older versions
                ForEach ($oldModule in $installedModules) {
                    If ($oldModule.Version -ne $latestModule.Version) {
                        Remove-Module -Name $module -Force -ErrorAction SilentlyContinue
                        Uninstall-Module -Name $module -RequiredVersion $oldModule.Version -Force -ErrorAction Stop
                        Write-Output "Removed older version $($oldModule.Version) of $module."
                    }
                }
            }
            Catch {
                If ($logFile) {
                    Add-Content -Path $logFile -Value "ERROR: Failed to update $module. $_"
                }
                If (Show-Error) {
                    Show-Error -errorMsg "Error: Failed to update module. $_"
                }
                Write-Warning "Failed to update $module. Error: $_"
                Exit
            }
        }
        
        Import-Module $module -Force
    }
}