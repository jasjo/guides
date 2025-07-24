<#
.SYNOPSIS
    The script adds and updates Message Center Posts in a SharePoint Online list

.DESCRIPTION
    This script is to be used along with the created guide to create a Microsoft Message Center SharePoint Online list that can be used for tracking organizational change.
#>

# Connect to a SharePoint Online site using certificate thumbprint
# Replace siteUrl, clientId, thumbprint, list name, and tenant with values from your tenant and app registration

$siteUrl = "https://jasjocom.sharepoint.com/sites/Test1"
$clientId = "2d167401-2514-4ce3-a6bc-b235028c1632"
$thumbprint = "675E3BC477F78FFBB32F07463135334FAD728328"
$tenant = "jasjocom.onmicrosoft.com"
$listName = "MC Test 3"

Connect-PnPOnline -Url $siteUrl -ClientId $clientId -Thumbprint $thumbprint -Tenant $tenant

$spoListItems = Get-Pnplistitem -list $listName
$messages = Get-PnPMessageCenterAnnouncement

foreach ($message in $messages) {
    if ($spoListItems.FieldValues.Title -contains $message.Id) {
        $existingSpoListItem = $spoListItems | ? { $_.FieldValues.Title -eq $message.id }
        
        if ($existingSpoListItem.FieldValues.LastModifiedDate -ne $message.LastModifiedDateTime.ToString("MM-dd-yyyy")) {
            Write-Host "MessageId $($message.id) already found. Updating list."
                      
            Set-PnPListItem -List $listName -Identity $existingSpoListItem.Id -Values @{
                "Title"            = $message.id; 
                "Category"         = $message.category.ToString(); 
                "StartTime"        = $message.startDateTime;
                "EndTime"          = $message.endDateTime;
                "Services"         = ($message.Services -join ";#").ToString();
                "PostTitle"        = $message.title;
                "Post"             = $message.body.content;
                "Tags"             = ($message.tags -join ";#");
                "Severity"         = $message.severity.toSTring();
                "IsMajorChange"    = $message.isMajorChange;
                "ActByDate"        = $message.ActionRequiredByDateTime;
                "LastModifiedDate" = $message.LastModifiedDateTime.ToString("MM-dd-yyyy");                
            }
        }
        else {
            Write-Host "MessageId $($message.id) No changes found."
        }

    }
    else {
        Write-Host "MessageId $($message.id) not found. Adding to list."  
       
        Add-PnPListItem -List $listName -Values @{
            "Title"            = $message.id; 
            "Category"         = $message.category.ToString(); 
            "StartTime"        = $message.startDateTime;
            "EndTime"          = $message.endDateTime;
            "Services"         = ($message.Services -join ";#");
            "PostTitle"        = $message.title;
            "Post"             = $message.body.content;
            "Tags"             = ($message.tags -join ";#");
            "Severity"         = $message.severity.ToString();
            "IsMajorChange"    = $message.isMajorChange;
            "ActByDate"        = $message.ActionRequiredByDateTime;
            "LastModifiedDate" = $message.LastModifiedDateTime.ToString("MM-dd-yyyy");
            "Status"           = "New"
        }
    }
}
Disconnect-PnPOnline
