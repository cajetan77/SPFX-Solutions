<#
.SYNOPSIS
    Creates the Desk Master and Desk Booking SharePoint lists for the Desk Booking SPFx solution.

.DESCRIPTION
    Requires PnP.PowerShell. Run while connected to your target SharePoint site:
    Connect-PnPOnline -Url "https://contoso.sharepoint.com/sites/DeskBooking" -Interactive

.PARAMETER SiteUrl
    SharePoint site URL where the lists should be created.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$SiteUrl = "https://caje77sharepoint.sharepoint.com/sites/App-3"
)

$ErrorActionPreference = "Stop"

function Ensure-Connected {
    param([string]$Url)

   
    Connect-PnPOnline -Url $Url -ClientId "dc223b11-5ab5-4a33-988a-3474b25eb9be" -Thumbprint
    
}

Ensure-Connected -Url $SiteUrl

$deskMasterTitle = "Desk Master"
$deskBookingTitle = "Desk Booking"

if (-not (Get-PnPList -Identity $deskMasterTitle -ErrorAction SilentlyContinue)) {
    $deskMasterList = New-PnPList -Title $deskMasterTitle -Template GenericList -Url "Lists/DeskMaster"
    Add-PnPField -List $deskMasterList -DisplayName "Location" -InternalName "Location" -Type Text -AddToDefaultView
    Add-PnPField -List $deskMasterList -DisplayName "Is Active" -InternalName "IsActive" -Type Boolean -AddToDefaultView
    Set-PnPField -List $deskMasterList -Identity "Title" -Values @{ Title = "Desk Name" }

    Add-PnPListItem -List $deskMasterList -Values @{
        Title    = "Desk A1"
        Location = "Floor 1 - Open Area"
        IsActive = $true
    }
    Add-PnPListItem -List $deskMasterList -Values @{
        Title    = "Desk B2"
        Location = "Floor 2 - Window Row"
        IsActive = $true
    }

    Write-Host "Created '$deskMasterTitle' list with sample desks." -ForegroundColor Green
}
else {
    Write-Host "'$deskMasterTitle' already exists." -ForegroundColor Yellow
}

$deskMasterList = Get-PnPList -Identity $deskMasterTitle

if (-not (Get-PnPList -Identity $deskBookingTitle -ErrorAction SilentlyContinue)) {
    $deskBookingList = New-PnPList -Title $deskBookingTitle -Template GenericList -Url "Lists/DeskBooking"
    # Add-PnPField -List $deskBookingList -DisplayName "Desk" -InternalName "Desk" -Type Lookup -LookupList $deskMasterList.Id -LookupField "Title" -Required -AddToDefaultView
    Add-PnPField -List $deskBookingList -DisplayName "Booking Date" -InternalName "BookingDate" -Type DateTime -Required -AddToDefaultView
    Set-PnPField -List $deskBookingList -Identity "BookingDate" -Values @{ DisplayFormat = 0; FriendlyDisplayFormat = 0 }
    Add-PnPField -List $deskBookingList -DisplayName "Start Time" -InternalName "StartTime" -Type Text -Required -AddToDefaultView
    Add-PnPField -List $deskBookingList -DisplayName "End Time" -InternalName "EndTime" -Type Text -Required -AddToDefaultView
    Add-PnPField -List $deskBookingList -DisplayName "Booker (Internal)" -InternalName "BookerPerson" -Type User -AddToDefaultView
    Add-PnPField -List $deskBookingList -DisplayName "Booker Name" -InternalName "BookerName" -Type Text -Required -AddToDefaultView
    Add-PnPField -List $deskBookingList -DisplayName "Booker Email" -InternalName "BookerEmail" -Type Text -Required -AddToDefaultView
    Add-PnPField -List $deskBookingList -DisplayName "Booking Status" -InternalName "BookingStatus" -Type Choice -Choices "Booked", "Cancelled" -Required -AddToDefaultView

    Write-Host "Created '$deskBookingTitle' list." -ForegroundColor Green
}
else {
    Write-Host "'$deskBookingTitle' already exists." -ForegroundColor Yellow
}

Write-Host "Provisioning complete." -ForegroundColor Green
