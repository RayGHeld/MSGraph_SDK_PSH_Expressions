# This script uses the Microsoft Graph Powershell Module: https://docs.microsoft.com/en-us/graph/powershell/get-started
# please ensure the module is installed.
 
$outFilePath = 'c:\temp\lastSignIns.csv'
$hasError = $false
 
Connect-MgGraph -scopes "Directory.Read.All", "AuditLog.Read.All"
Select-MgProfile -Name beta

try{
    $users = Get-MgUser -Top 10 -Property "id, userPrincipalName, displayName, otherMails, signInActivity, onPremisesExtensionAttributes" `
        -ConsistencyLevel eventual -Count $count `
        -ErrorAction Stop -ErrorVariable GraphError | `
        Select-Object -Property id, userPrincipalName, displayName, `
        @{Name='otherEmails'                      ; Expression={@($_.otherMails)[0]};}, `
        @{Name='lastSignInDateTime'               ; Expression={$_.signInActivity.lastSignInDateTime};}, `
        @{Name='lastNonInteractiveSignInDateTime' ; Expression={$_.signInActivity.lastNonInteractiveSignInDateTime};}, `
        @{Name='extensionAttribute14'             ; Expression={$_.onPremisesExtensionAttributes.extensionAttribute14};} `

    
    Write-Host "Downloaded report to memory..."
} catch {
    Write-Host "Error downloading report: $GraphError.Message"
    $hasError = $true
}

if(!$hasError){
    try{
        Write-Output $users | Format-table -Property *
        Write-Host "Writing to .csv file..."
        $users | Export-Csv -Path $outFilePath
        Write-Host "Report saved at $outFilePath"
    } catch {
        Write-Host "Error saving .csv: $_.ErrorDetails.Message"
    }
}
 
# Disconnect-MgGraph
