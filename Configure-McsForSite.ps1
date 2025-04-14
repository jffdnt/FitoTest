param(
    [Parameter(Mandatory=$true)]
    [string]$siteUrl,
    [Parameter(Mandatory=$true)]
    [string]$botUrl,
    [string]$botName = "Copilot",
    [bool]$greet = $false,
    [Parameter(Mandatory=$true)]
    [string]$customScope,
    [Parameter(Mandatory=$true)]
    [string]$clientId,
    [Parameter(Mandatory=$true)]
    [string]$authority,
    [string]$buttonLabel = "Chat",
    [bool]$useFiToTemplate = $false,
    [string]$botAvatarImage = "images/fito.png"
)

# Connect to the site
Connect-PnPOnline -Url $siteUrl -Interactive

# Get the user custom action
$userCustomAction = Get-PnPCustomAction -Scope Site | Where-Object { $_.ClientSideComponentId -eq "[YOUR_COMPONENT_ID]" }

if ($userCustomAction) {
    # Create the JSON properties
    $jsonProperties = @{
        botURL = $botUrl
        botName = $botName
        greet = $greet
        customScope = $customScope
        clientID = $clientId
        authority = $authority
        buttonLabel = $buttonLabel
        useFiToTemplate = $useFiToTemplate
        botAvatarImage = $botAvatarImage
    }

    # Convert to JSON and escape quotes
    $jsonPropertiesString = $jsonProperties | ConvertTo-Json -Compress
    $jsonPropertiesString = $jsonPropertiesString.Replace('"', '&quot;')

    # Update the user custom action
    Set-PnPCustomAction -Identity $userCustomAction.Id -ClientSideComponentProperties $jsonPropertiesString

    Write-Host "Custom action updated successfully." -ForegroundColor Green
} else {
    Write-Host "Custom action not found." -ForegroundColor Red
}