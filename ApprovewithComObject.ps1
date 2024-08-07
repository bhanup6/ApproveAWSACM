# Create Outlook COM object
$urlPattern = "https:\/\/[^\s]+"
$Outlook = New-Object -ComObject Outlook.Application
$namespace = $Outlook.GetNameSpace("MAPI")
$inbox = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox)

# Filter for unread emails
$unreadEmails = $inbox.Items.Restrict("[Unread] = True")

# Process unread emails
foreach ($email in $unreadEmails) {
    
    if (($email.Body -match 'To approve this request, go to Amazon Certificate Approvals') -and ($email.Body -match $urlPattern)) {
    $url =$matches[0]
    $updatedurl=$url.Trim() -replace "\\r\\n","" 
        Write-Host $updatedurl
        $res=Invoke-RestMethod -Uri $updatedurl -Method Get -UseDefaultCredentials 
            $html =$res
            $formAction = ($html -split 'form action="')[1].Split('"')[0]
            $validationToken = ($html -split 'name="validationToken" value="')[1].Split('"')[0]
            $validationArn = ($html -split 'name="validationArn" value="')[1].Split('"')[0]
            $domainName = ($html -split 'name="domainName" value="')[1].Split('"')[0]
            $accountId = ($html -split 'name="accountId" value="')[1].Split('"')[0]
            $region = ($html -split 'name="region" value="')[1].Split('"')[0]
            $certificateIdentifier = ($html -split 'name="certificateIdentifier" value="')[1].Split('"')[0]
            $formData = @{
                validationToken = $validationToken
                validationArn = $validationArn
                domainName = $domainName
                accountId = $accountId
                region = $region
                certificateIdentifier = $certificateIdentifier
                validationApprovalStatus = "APPROVED"
            }
            $formdata|FT
            try {
                    $response = Invoke-WebRequest -Uri $formAction -Method Post -Body $formData
                     if ($response.StatusCode -eq 200) {
                             Write-Host "Certificate approval successful!"
                             Update-MgUserMessage -UserId $userId -MessageId $m.Id -BodyParameter $params
                             Write-Host "Email marked as read successfully."
                }    else{
                             Write-Host "Certificate approval failed. Status code: $($response.StatusCode)"
                    }
}          catch {
                     Write-Host "An error occurred: $_"
}
    
           }
    }

    
    
    # Optionally, mark the email as read
    # $email.UnRead = $false
    # $email.Save()

# Release COM object
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null