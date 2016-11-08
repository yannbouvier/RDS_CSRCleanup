# Init vars
$CSRCurrentKey = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Providers\Client Side Rendering Print Provider"
$CSRBackupKey = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Providers\Client Side Rendering Print Provider.bak"
$CSRBackupKeyName = "Client Side Rendering Print Provider.bak"
$logFile = "C:\Scripts\CSRCleanup\CSRCleanup.log"

# Cleaning screen
Clear-Host

# Stopping spooler service
$dateNow = get-date -uformat "%d/%m/%Y-%H:%M:%S"
Add-Content -Path $logFile -Value "$dateNow - PROCESS : Stopping spooler service..."
#Write-Host "PROCESS : Stopping spooler service..." -ForegroundColor Yellow

Stop-Service -Name Spooler
$dateNow = get-date -uformat "%d/%m/%Y-%H:%M:%S"
If((Get-Service | where {$_.Name -eq "Spooler"}).Status -eq "Running")
    {
        Add-Content -Path $logFile -Value "$dateNow - ERROR : Stopping spooler service => Process ABORTED"
        #Write-Host "ERROR : Stopping spooler service => Process ABORTED" -ForegroundColor Red
    }
Else
    {
    Add-Content -Path $logFile -Value "$dateNow - COMPLETE : Stopping spooler service"
    #Write-Host "COMPLETE : Stopping spooler service" -ForegroundColor Green

    # Check previous backup key then delete if exist
    If(Test-Path $CSRBackupKey)
        {
        $dateNow = get-date -uformat "%d/%m/%Y-%H:%M:%S"
        Add-Content -Path $logFile -Value "$dateNow - PROCESS : Removing CSR Backup Key..."
        #Write-Host "PROCESS : Removing CSR Backup Key..." -ForegroundColor Yellow
        Remove-Item $CSRBackupKey -Recurse -Force 2> $null
        }

    # Test operation result
    $dateNow = get-date -uformat "%d/%m/%Y-%H:%M:%S"
    If(Test-Path $CSRBackupKey)
        {
        Add-Content -Path $logFile -Value "$dateNow - ERROR : Removing CSR Backup Key => Process ABORTED"
        #Write-Host "ERROR : Removing CSR Backup Key => Process ABORTED" -ForegroundColor Red
        }
    Else
        {
            Add-Content -Path $logFile -Value "$dateNow - COMPLETE : Removing CSR Backup Key"
            #Write-Host "COMPLETE : Removing CSR Backup Key" -ForegroundColor Green
            Add-Content -Path $logFile -Value "$dateNow - PROCESS : Create backup Key..."
            #Write-Host "PROCESS : Create backup Key..." -ForegroundColor Yellow

            # Create backup Key
            Rename-Item $CSRCurrentKey -NewName $CSRBackupKeyName 2> $null

            # Test operation result
            $dateNow = get-date -uformat "%d/%m/%Y-%H:%M:%S"
            If(!(Test-Path $CSRBackupKey))
                {
                Add-Content -Path $logFile -Value "$dateNow - ERROR : Create backup Key => Process ABORTED"
                #Write-Host "ERROR : Create backup Key => Process ABORTED" -ForegroundColor Red
                }
            Else
                {
                # Removing Current Key
                Remove-Item $CSRCurrentKey -Recurse -Force 2> $null

                # Test operation result
                $dateNow = get-date -uformat "%d/%m/%Y-%H:%M:%S"
                If(!(Test-Path $CSRCurrentKey))
                    {
                    Add-Content -Path $logFile -Value "$dateNow - COMPLETE : Create backup Key"
                    #Write-Host "COMPLETE : Create backup Key" -ForegroundColor Green

                    # Creating new empty key
                    $dateNow = get-date -uformat "%d/%m/%Y-%H:%M:%S"
                    Add-Content -Path $logFile -Value "$dateNow - PROCESS : Creating new CSR key"
                    #Write-Host "PROCESS : Creating new CSR key" -ForegroundColor Yellow
                    New-Item $CSRCurrentKey 2> $null

                    $dateNow = get-date -uformat "%d/%m/%Y-%H:%M:%S"
                    # Test operation result
                    If((!(Test-Path $CSRCurrentKey)))
                        {
                        Add-Content -Path $logFile -Value "$dateNow - ERROR : Creating new CSR key - ALERT"
                        #Write-Host "ERROR : Creating new CSR key - ALERT" -ForegroundColor Red
                        }
                    Else
                        {
                        Add-Content -Path $logFile -Value "$dateNow - COMPLETE : Creating new CSR key"
                        #Write-Host "COMPLETE : Creating new CSR key" -ForegroundColor Green
                        }
                    }
                Else
                    {
                    Add-Content -Path $logFile -Value "$dateNow - COMPLETE : Creating new CSR key"
                    #Write-Host "COMPLETE : Creating new CSR key" -ForegroundColor Red
                    }
                }
        }

    # Starting spooler service
    $dateNow = get-date -uformat "%d/%m/%Y-%H:%M:%S"
    Add-Content -Path $logFile -Value "$dateNow - PROCESS : Starting spooler service..."
    #Write-Host "PROCESS : Starting spooler service..." -ForegroundColor Yellow
    Start-Service -Name Spooler
    $dateNow = get-date -uformat "%d/%m/%Y-%H:%M:%S"
    If((Get-Service | where {$_.Name -eq "Spooler"}).Status -eq "Running")
        {
        Add-Content -Path $logFile -Value "$dateNow - COMPLETE : Starting spooler service"
        #Write-Host "COMPLETE : Starting spooler service" -ForegroundColor Green
        }
    Else
        {
        Add-Content -Path $logFile -Value "$dateNow - ERROR : Starting spooler service - ALERT"
        #Write-Host "ERROR : Starting spooler service - ALERT" -ForegroundColor Red
        }
}
Add-Content -Path $logFile -Value "##############################################################"
