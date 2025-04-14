# AD Server Audit
<#
The object is to generate an audit of servers in AD. This script does an ICMP 
and WinRM service test to see if the server is 'alive'. It gets the creation 
date and last change date; useful to know if the server needs to be removed 
from AD.
#>

param (
    [Parameter()]
    [string]$ExcelFile = "$([environment]::getfolderpath("mydocuments"))\AD Server Audit.xlsx"

)

# variables
$serverListHash = @()
$serverHash = [ordered]@{}
$counter = 0

# query AD servers and get all properties
$serversAD = Get-ADComputer -Filter "OperatingSystem -like '*Server*'" -Properties *

<#
itterate through serversAD
add wanted properties to hash table
add hash table to list
#>
foreach ($server in $serversAD) {
    # increase counter
    $counter++
    # server name
    $serverHash.Add("Server", "$($server.Name)")
    # type of OS
    $serverHash.Add("OS", "$($server.OperatingSystem)")
    # directory tree path
    $serverHash.Add("AD Tree", "$($server.CanonicalName)")
    # creation date
    $serverHash.Add("Created", "$($server.whenCreated)")
    # last modified
    $serverHash.Add("Last Changed", "$($server.whenChanged)")

    # ICMP
    try {
        # run connection test
        $connTest = Test-Connection -ComputerName $server.Name -ErrorAction Stop -Count 1
    } catch {
        $connTest = $_
    } finally {
        # add up down status to hash
        if (!$connTest.Exception.Message) {
            $serverHash.Add("ICMP","up")
        } else {
            $serverHash.Add("ICMP","down")
        }
    }

    # WinRM
    try {
        # run WinRM test
        $wsmanTest = Test-WSMan -ComputerName $server.Name -ErrorAction Stop
    } catch {
        $wsmanTest = $_
    } finally {
        # add up down status to hash, if successful gather server information
        if (!$wsmanTest.Exception.Message) {
            $serverHash.Add("WinRM","up")
            # gather server information
            try {
                # gather system information
                $cimData = Get-CimInstance -ClassName CIM_ComputerSystem -ComputerName $server.Name -ErrorAction Stop
            } catch {
                $cimData = $_
            } finally {
                # add physical and logical cpu count, add hardware model
                if (!$cimData.Exception.Message) {
                    $serverHash.add("Physical Proc Count", "$($cimData.NumberOfProcessors)")
                    $serverHash.add("Logical Proc Count", "$($cimData.NumberOfLogicalProcessors)")
                    $serverHash.add("Model", "$($cimData.Model)")
                } else {
                    # give error code for failure reason
                    Write-Host "$($server.Name) failed Get-CimInstance." -ForegroundColor Yellow
                    Write-Host "$($cimData.Exception.Message)" -ForegroundColor Red
                }
            }
        } else {
            $serverHash.Add("WinRM","down")
        }
    }


    # add hash table to server list
    $serverListHash += [pscustomobject]$serverHash

    # Rest serverHash var for next itteration
    $serverHash = [ordered]@{}
    # write progress to screen
    Write-Progress -Activity "Gathering Data" -Status "Processing $counter of $($serversAD.Length)" -PercentComplete $(($counter/$serversAD.Length) * 100)
}

# export data to excel
try {
    # export gathered data to excel file
    $writeFile = $serverListHash | Export-Excel -Path $ExcelFile -FreezeTopRow -WorksheetName "AD Server Audit"
} catch {
    $writeFile = $_
} finally {
    # check of file write failed
    if (!$writeFile.Exception.Message) {
        Write-Progress -Activity "Gathering Data" -Status "Done" -Completed
        Write-Host "Wrote to file: $ExcelFile"
    } else {
        # give error 
        Write-Host $writeFile
    }
}
