param(

    [int]$days = 7, #Days worth of logs
    [string]$singleCompany = $False, # $True when running against only one company
    [string]$logPath = "$PSScriptRoot\Logs\GatherData-Log - $date.txt",
    [string]$dataPath = "$PSScriptRoot\Raw Data\auditdata-$($date)-logins.csv"
)

function write-log {param([string]$message)
    Write-host "`n$(get-date): $message" 
    Write-Output "$(get-date): $message" | Out-File -FilePath $logPath -Append
      
  }

################################## Configuration Section ##################################
# Accounts file
$accounts = import-csv -path "$PSScriptRoot\Accounts.csv"

#retries and timeout
$retryMax = 3
$retrySleepMinutes = 2

# Date for file names
$date = (Get-Date -Format dd-MM-yy)

$auditData = @()

####################################### Section End #######################################

if ($singleCompany -eq $true) {
    $companies = $accounts | ForEach-Object { $_.name }
    Write-host "Beginning single company Audit Search`n" -ForegroundColor Cyan
    Write-Host "Please Select Company From List below `n" -ForegroundColor Yellow
    
    $i = 0
    ForEach ($company in $companies) {
        Write-Host "$i : $company "
        $i++
    }

    $companyNum = Read-Host "`nEnter Number: "
    $accounts = $accounts | Where-Object name -EQ $companies[$companyNum] | Select-Object *
}

$accounts | ForEach-Object {
    $companyName = $_.name
    write-log -message "############## $companyName ##############"

    $cred = New-Object System.Management.Automation.PSCredential `
        -ArgumentList $_.email, $(ConvertTo-SecureString $_.password)
 
    $session = New-PSSession -ConfigurationName Microsoft.Exchange `
        -ConnectionUri https://outlook.office365.com/powershell-liveid/ `
        -Credential $cred -Authentication Basic -AllowRedirection    
    Import-PSSession $session 

    $AuditLogEnabled = (Get-AdminAuditLogConfig).UnifiedAuditLogIngestionEnabled 
    if ($AuditLogEnabled -ne $true) { 
        write-log -message "Attempting to enable auditing" 
        try { Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true -erroraction stop }
        catch { write-log -message "Enabling Auditing failed" } 
    }

    $LogSearchParams = @{
        enddate        = $(get-date)
        StartDate      = $((get-date).AddDays(-$days))
        ResultSize     = 1000
        sessionID      = $([DateTime]::Now.ToString().Replace('/', '_')).ToString
        sessionCommand = 'ReturnNextPreviewPage'
        RecordType     = 'AzureActiveDirectoryStsLogon'
        erroraction    = 'stop'
    }

    $retryCount = 0  # Create retry counter
    while ($true) {
       
        try {
            [array]$results = Search-UnifiedAuditLog @LogSearchParams }
        catch { write-log -message "Unable to search audit data" }
        
        Start-Sleep -Seconds 10

        # if results contains data... 
        if ( $results -ne $null -or $results.Count -ne 0 ) {   
            $auditData += $results
            $results = $null
            Write-log -message "Audit Data found"    
            continue 
        }
        # if no data found within results...
        if ($results -eq $null -or $results.Count -eq 0 ) {
        
            write-log -message "No data found begining sleep"     
            Start-Sleep -Seconds $(60 * $retrySleepMinutes) 

            # if results contains data after sleep... 
            if ( $results -ne $null -or $results.Count -ne 0 ) {
                $auditData += $results # Store data in auditData array
                $results = $null
                write-log -message "Audit Data found" 
                continue # search for more data
            }
            else {
                $retryCount ++ # Increment retry count
            }
        }
        # max amount of retries exceeded stop searching...
        if ($retryCount -gt $retryMax -or $retryCount -eq $retryMax ) {   
            write-log -message "Search complete"
            # Stop searching
            break 
        }           
    }
    
    write-log -message "Converting and sorting data from json"
    $logArray = @()

    # Sort data and convert auditData from Json
    $auditData | sort -Property UserIds | %{$_.auditdata} | ConvertFrom-Json | select -Property UserID,Operation,CreationTime,ClientIP `
    |Where-Object Operation -eq 'UserLoggedIn'`
    |ForEach-Object { 

        $userLog = [PSCustomObject]@{
            Company = $companyNameue
            User = $_.userID
            Operation = $_.Operation
            CreationTime = $_.CreationTime
            ClientIP = $_.ClientIP
        }     
        $logArray += $userLog
    }
    write-log -message "convert and sort complete"
    # add data for each company to a complete array
    $completeData += $logArray 
    
    write-log -message "closing current company connection"
    Get-PSSession | Remove-PSSession 
}
write-log -message "All companies processed, exporting data"
$completeData | Export-Csv -Path $dataPath


