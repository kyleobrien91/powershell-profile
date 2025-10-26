# Ensures that the script is being run with administrative privileges.
# This is necessary because the script will be making changes to the system,
# such as installing software and modifying the user's PowerShell profile.
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "This script requires administrative privileges. Please run it as an Administrator."
    break
}

# Defines a function to check for an active internet connection.
# This is crucial because the script needs to download various files and packages from the internet.
# It works by attempting to send a single ping to www.google.com.
function Test-InternetConnection {
    try {
        Test-Connection -ComputerName www.google.com -Count 1 -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        Write-Warning "An internet connection is required to proceed, but one could not be established. Please check your network connection."
        return $false
    }
}

# Defines a function to download and install Nerd Fonts.
# Nerd Fonts are a collection of popular programming fonts that have been patched with a large number of icons and symbols.
# These are often used with tools like Oh My Posh to create a more visually appealing and informative command prompt.
function Install-NerdFonts {
    param (
        [string]$FontName = "CascadiaCode",
        [string]$FontDisplayName = "CaskaydiaCove NF",
        [string]$Version = "3.2.1"
    )

    try {
        # Checks if the specified font is already installed on the system.
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
        $fontFamilies = (New-Object System.Drawing.Text.InstalledFontCollection).Families.Name
        if ($fontFamilies -notcontains "${FontDisplayName}") {
            # If the font is not installed, it is downloaded from the official Nerd Fonts GitHub repository.
            $fontZipUrl = "https://github.com/ryanoasis/nerd-fonts/releases/download/v${Version}/${FontName}.zip"
            $zipFilePath = "$env:TEMP\${FontName}.zip"
            $extractPath = "$env:TEMP\${FontName}"

            $webClient = New-Object System.Net.WebClient
            $webClient.DownloadFileAsync((New-Object System.Uri($fontZipUrl)), $zipFilePath)

            # Waits for the download to complete before proceeding.
            while ($webClient.IsBusy) {
                Start-Sleep -Seconds 2
            }

            # Extracts the downloaded font files and copies them to the system's Fonts directory.
            Expand-Archive -Path $zipFilePath -DestinationPath $extractPath -Force
            $destination = (New-Object -ComObject Shell.Application).Namespace(0x14)
            Get-ChildItem -Path $extractPath -Recurse -Filter "*.ttf" | ForEach-Object {
                If (-not(Test-Path "C:\Windows\Fonts\$($_.Name)")) {
                    $destination.CopyHere($_.FullName, 0x10)
                }
            }
            # Removes the temporary files that were created during the installation process.
            Remove-Item -Path $extractPath -Recurse -Force
            Remove-Item -Path $zipFilePath -Force
        } else {
            Write-Host "The ${FontDisplayName} font is already installed."
        }
    }
    catch {
        Write-Error "An error occurred while trying to download or install the ${FontDisplayName} font: $_"
    }
}

# Before proceeding with the rest of the script, it's important to ensure that there is an active internet connection.
if (-not (Test-InternetConnection)) {
    break
}

# This section of the script handles the creation or updating of the user's PowerShell profile.
# The PowerShell profile is a script that runs every time a new PowerShell session is started.
# It can be used to customize the shell's environment, such as by adding aliases, functions, and modules.
if (!(Test-Path -Path $PROFILE -PathType Leaf)) {
    try {
        # Detects the version of PowerShell being used and creates the appropriate profile directories if they do not already exist.
        $profilePath = ""
        if ($PSVersionTable.PSEdition -eq "Core") {
            $profilePath = "$env:userprofile\Documents\Powershell"
        }
        elseif ($PSVersionTable.PSEdition -eq "Desktop") {
            $profilePath = "$env:userprofile\Documents\WindowsPowerShell"
        }

        if (!(Test-Path -Path $profilePath)) {
            New-Item -Path $profilePath -ItemType "directory"
        }

        # Downloads a pre-configured PowerShell profile from a GitHub repository and saves it to the user's profile path.
        Invoke-RestMethod https://github.com/ChrisTitusTech/powershell-profile/raw/main/Microsoft.PowerShell_profile.ps1 -OutFile $PROFILE
        Write-Host "Your PowerShell profile has been created at [$PROFILE]."
        Write-Host "If you wish to make any personal customizations, please do so in the file located at [$profilePath\Profile.ps1]. This is important because the installed profile includes an updater that will overwrite any changes made directly to the main profile file."
    }
    catch {
        Write-Error "An error occurred while creating or updating your PowerShell profile: $_"
    }
}
else {
    try {
        # If a PowerShell profile already exists, it is backed up before being replaced with the new one.
        $backupPath = Join-Path (Split-Path $PROFILE) "oldprofile.ps1"
        Move-Item -Path $PROFILE -Destination $backupPath -Force
        Invoke-RestMethod https://github.com/ChrisTitusTech/powershell-profile/raw/main/Microsoft.PowerShell_profile.ps1 -OutFile $PROFILE
        Write-Host "✅ Your PowerShell profile at [$PROFILE] has been successfully updated."
        Write-Host "📦 Your old profile has been backed up to [$backupPath]."
        Write-Host "⚠️ Please note that any personal customizations from your old profile should be moved to [$HOME\Documents\PowerShell\Profile.ps1] to avoid being overwritten by the profile's built-in updater."
    }
    catch {
        Write-Error "❌ An error occurred while backing up and updating your PowerShell profile: $_"
    }
}

# This section of the script installs Oh My Posh, a popular tool for customizing the appearance of the PowerShell prompt.
try {
    winget install -e --accept-source-agreements --accept-package-agreements JanDeDobbeleer.OhMyPosh
}
catch {
    Write-Error "An error occurred while installing Oh My Posh: $_"
}

# This section of the script installs the Cascadia Code Nerd Font, which is a popular choice for use with Oh My Posh.
Install-NerdFonts -FontName "CascadiaCode" -FontDisplayName "CaskaydiaCove NF"

# This section of the script performs a final check to ensure that all of the necessary components have been installed correctly.
# It then displays a message to the user, letting them know whether the setup was successful.
if ((Test-Path -Path $PROFILE) -and (winget list --name "OhMyPosh" -e) -and ($fontFamilies -contains "CaskaydiaCove NF")) {
    Write-Host "The setup has been completed successfully. Please restart your PowerShell session to see the changes."
} else {
    Write-Warning "The setup has been completed, but some errors were encountered. Please review the error messages above."
}

# This section of the script installs Chocolatey, a popular package manager for Windows.
try {
    Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
}
catch {
    Write-Error "An error occurred while installing Chocolatey: $_"
}

# This section of the script installs the Terminal-Icons module, which provides a set of icons that can be used to identify different types of files in the terminal.
try {
    Install-Module -Name Terminal-Icons -Repository PSGallery -Force
}
catch {
    Write-Error "An error occurred while installing the Terminal Icons module: $_"
}
# This section of the script installs zoxide, a tool that provides a more intelligent `cd` command.
try {
    winget install -e --id ajeetdsouza.zoxide
    Write-Host "zoxide has been installed successfully."
}
catch {
    Write-Error "An error occurred while installing zoxide: $_"
}
