# ========================
# Step 1: Get usage data as JSON
# ========================
$reportUrl = "https://graph.microsoft.com/beta/reports/getMicrosoft365CopilotUsageUserDetail(period='$period')"
$response = Invoke-RestMethod -Uri $reportUrl -Headers @{ Authorization = "Bearer $($token.AccessToken)" }
$usageData = $response.value

# ========================
# Step 2: Get all users and departments
# ========================
$users = @()
$nextLink = 'https://graph.microsoft.com/v1.0/users?$select=displayName,userPrincipalName,department&$top=999'
do {
    $resp = Invoke-RestMethod -Uri $nextLink -Headers @{ Authorization = "Bearer $accessToken" }
    $users += $resp.value
    $nextLink = $resp.'@odata.nextLink'
} while ($nextLink)

# ========================
# Step 3: Map userPrincipalName → department
# ========================
$userDeptMap = @{}
foreach ($u in $users) {
    $userDeptMap[$u.userPrincipalName.ToLower()] = $u.department
}

# ========================
# Step 4: Join usage with departments
# ========================
$joined = $usageData | ForEach-Object {
    $dept = $userDeptMap[$_.userPrincipalName.ToLower()]
    [PSCustomObject]@{
        UserPrincipalName     = $_.userPrincipalName
        Department            = $dept
        LastActivityDate      = $_.lastActivityDate
        CopilotChatLastDate   = $_.copilotChatLastActivityDate
        Product               = $_.product
    }
}
# Logging the joined data
# ========================
$joined | Format-Table UserPrincipalName, Product, LastActivityDate, Department

# ========================
# ========================
# Step 5: Count Copilot Chat usage per department
# ========================
$grouped = $joined | Where-Object {
    $_.Product -eq "Copilot Chat" -and $_.CopilotChatLastDate
} | Group-Object Department | Sort-Object Count -Descending

# ========================
# Display results
# ========================
$grouped | Format-Table Name, Count