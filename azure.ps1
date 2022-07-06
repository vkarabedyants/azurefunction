using namespace System.Net
# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

$ErrorActionPreference = "Stop"
# Secrets are taken from Function App>>Configuration>>Application Settings.
# They are necessary to get id_token
$azureTenantId = $Env:AzureTenantId
$clientId = $Env:ApplicationClientId
$clientSecret = $Env:AzureClientSecret
$username = $Env:AccessTokenUserName
$password = $Env:AccessTokenPassword

# Parse incoming request body to get "Documaster_api_url", "displayName" and "employeeId" of an AD user
# $documasterAuthHeaderName = "X-LogicApp-Authorization"
# $documasterBearerToken = $Request.Headers[$documasterAuthHeaderName]
$requestBody = $Request.Body
$documasterApiUrl = $requestBody["Documaster_api_url"]
$documasterEmployeeItem = $requestBody["items"][0]
$documasterUserDisplayName = $documasterEmployeeItem["displayName"]
$documasterEmployeeId = $documasterEmployeeItem["employeeId"]
# From the provided DM (Documaster) documentation: 
# "The tag name will be generated based on the employee ID of the user in AD. The tag format is a combination of “employee ID” and “display name” of the user in AD."
# Create an expected DM tag title based on the description above
$expectedDocumasterTitle = "$documasterUserDisplayName [$documasterEmployeeId]"

# Write the obtained data to console (just for information/debugging purposes)
Write-Host "------------------------------------------Documaster_api_url: $documasterApiUrl"
Write-Host "---------------------------------------AD user's displayName: $documasterUserDisplayName"
Write-Host "----------------------------------------AD user's employeeId: $documasterEmployeeId"
Write-Host "----expected Documaster tag title (displayName [employeeId]): $expectedDocumasterTitle"


# Obtain Documaster id_token
$adAccessTokenApiUri = "https://login.microsoftonline.com/$azureTenantId/oauth2/v2.0/token"
$scope = "user.read openid profile offline_access"
$grantType = "password"
$contentType = "application/x-www-form-urlencoded"
$adAccessTokenBody = @{
    client_id = $clientId
    scope = $scope
    client_secret = $clientSecret
    grant_type = $grantType
    username = $username
    password=$password
}
$adAccessTokenResult = Invoke-RestMethod -Method Post -Body $adAccessTokenBody -Uri $adAccessTokenApiUri -ContentType $contentType
$adAccessToken = $adAccessTokenResult.id_token
#Write-Host "-----------adAccessTokenAnother: " $adAccessToken

# Lookup a tag by the externalId. Obtain "id" and "title" properties of the found tag.
$documasterHeaders = @{"Authorization" = "Bearer $adAccessToken"; "Content-Type" = "application/json" }
$documasterLookupTagUri = "$documasterApiUrl/tag/lookup"
$documasterLookupTagBody = @{"query" = "externalIds.externalId=@externalId"; "parameters" = @{"@externalId" = $documasterEmployeeId}} | ConvertTo-Json
$documasterLookupTagResult = Invoke-RestMethod -Method Post -Uri $documasterLookupTagUri -Headers $documasterHeaders -Body $documasterLookupTagBody
$documasterTagData = $($documasterLookupTagResult).data
$documasterTagId = $($documasterTagData).id
$documasterTagTitle = $($documasterTagData).title

# Write the obtained data to console (just for information/debugging purposes)
Write-Host "--------------documasterTagId: $documasterTagId"
Write-Host "-----------documasterTagTitle: $documasterTagTitle"
Write-Host "----expected Documaster Title: $expectedDocumasterTitle"

# Check if the "title" of the obtained DM tag (i.e. $documasterTagTitle) equals information from AD (i.e. $expectedDocumasterTitle)
if (($documasterTagTitle).ToLower() -ne ($expectedDocumasterTitle).ToLower()) {
    Write-Host "AD display name and DM tag title are different. Will update Documaster side..."
    # Update Documaster tag title
    $documasterUpdateTagApiUri = "$documasterApiUrl/tag/$documasterTagId"
    $documasterUpdateTagBody = @{data = @{title = $expectedDocumasterTitle}} | ConvertTo-Json
    $documasterUpdateTagResult = Invoke-RestMethod -Method Put -Uri $documasterUpdateTagApiUri -Headers $documasterHeaders -Body $documasterUpdateTagBody
    # Write-Host "result.data" $($documasterUpdateTagResult).data
} else {
    Write-Host " AD display name and DM tag title are the same...Nothing to change..."
}