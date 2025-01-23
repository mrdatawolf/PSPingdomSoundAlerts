# Function to play sound
function Play-Sound {
    param (
        [string]$soundPath
    )
    $extension = [System.IO.Path]::GetExtension($soundPath).ToLower()
    if ($extension -eq ".wav") {
        Write-Host "Playing WAV sound: $soundPath"
        $player = New-Object System.Media.SoundPlayer
        $player.SoundLocation = $soundPath
        $player.PlaySync()
    } elseif ($extension -eq ".mp3") {
        Write-Host "Playing MP3 sound: $soundPath"
        $player = New-Object -ComObject WMPlayer.OCX.7
        $player.URL = $soundPath
        $player.controls.play()
        Start-Sleep -Seconds $player.currentMedia.duration
        $player.close()
    } else {
        Write-Host "Unsupported sound file format: $soundPath"
    }
}

# Function to load .env file
function Load-EnvFile {
    param (
        [string]$envFilePath
    )
    Get-Content $envFilePath | ForEach-Object {
        if ($_ -match "^\s*([^#][^=]+)=(.*)\s*$") {
            $name = $matches[1].Trim()
            $value = $matches[2].Trim()
            [System.Environment]::SetEnvironmentVariable($name, $value)
        }
    }
}

# Load the .env file
$scriptPath = $MyInvocation.MyCommand.Path
$scriptDir = Split-Path -Path $scriptPath -Parent
$envFilePath = Join-Path -Path $scriptDir -ChildPath ".env"
Load-EnvFile -envFilePath $envFilePath
Write-Host "Loaded .env file from $envFilePath"

# Load Outlook COM object
$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")
$Inbox = $Namespace.GetDefaultFolder(6) # 6 corresponds to olFolderInbox
$PingdomFolder = $Inbox.Folders.Item($env:OUTLOOK_FOLDER)
Write-Host "Connected to Outlook and accessed folder: $env:OUTLOOK_FOLDER"

# Function to check emails
function Check-Emails {
    Write-Host "Checking for unread emails..."
    $unreadItems = $PingdomFolder.Items | Where-Object { $_.UnRead -eq $true }
    if ($unreadItems.Count -gt 0) {
        Write-Host "Found $($unreadItems.Count) unread email(s)"
        $latestEmail = $unreadItems | Sort-Object ReceivedTime -Descending | Select-Object -First 1
        Write-Host "Latest email subject: $($latestEmail.Subject)"
        foreach ($email in $unreadItems) {
            $email.UnRead = $false
            $email.Save()
        }
        if ($latestEmail.Subject -like "DOWN*") {
            Play-Sound -soundPath $env:DOWN_SOUND
        } elseif ($latestEmail.Subject -like "UP*") {
            Play-Sound -soundPath $env:UP_SOUND
        }
        Write-Host "Marked all unread emails as read"
    } else {
        Write-Host "No unread emails found"
    }
}

# Run the script in the background
while ($true) {
    Check-Emails
    Start-Sleep -Seconds 60
}