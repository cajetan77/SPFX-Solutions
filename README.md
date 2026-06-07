# Desk Booking SPFx Solution

SharePoint Framework web part for booking desks in SharePoint Online using two SharePoint lists as the backend.

## Features

- View available desks for a selected date and time range
- Book a desk with one click
- Cancel your own active bookings
- Weekday-only booking (Monday through Friday)
- Business hours restriction (7:00 AM to 6:00 PM)
- Double-booking prevention for overlapping time slots
- Clear status badges: Available, Booked, and Cancelled
- Validation messages and error handling

## SharePoint Lists

Create two lists on the target site before using the web part.

### Desk Master

| Column | Type | Internal name |
| --- | --- | --- |
| Desk Name | Single line of text (rename Title) | `Title` |
| Location | Single line of text | `Location` |
| Is Active | Yes/No | `IsActive` |

### Desk Booking

| Column | Type | Internal name |
| --- | --- | --- |
| Desk | Lookup to Desk Master | `Desk` |
| Booking Date | Date and time (date only) | `BookingDate` |
| Start Time | Single line of text (`HH:mm`) | `StartTime` |
| End Time | Single line of text (`HH:mm`) | `EndTime` |
| Booker Name | Single line of text | `BookerName` |
| Booker Email | Single line of text | `BookerEmail` |
| Booking Status | Choice (`Booked`, `Cancelled`) | `BookingStatus` |

Bookings are anonymous — users provide their name and email when booking. They are not linked to SharePoint sign-in.

### Automated provisioning

Run the included PnP PowerShell script against your site:

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
.\provisioning\Create-DeskBookingLists.ps1 -SiteUrl "https://yourtenant.sharepoint.com/sites/yoursite"
```

## Development

### Prerequisites

- Node.js 22.x
- SharePoint Online tenant for testing
- [PnP PowerShell](https://pnp.github.io/powershell/) for list provisioning

### Build and run

```powershell
npm install -g @rushstack/heft
npm install
heft start
```

### Package for deployment

```powershell
heft package-solution --production
```

Upload `sharepoint/solution/desk-booking.sppkg` to your app catalog, deploy the package, provision the lists, then add the **Desk Booking** web part to a page.

## Web part configuration

Use the property pane to set list titles if they differ from the defaults:

- Desk Master list title: `Desk Master`
- Desk Booking list title: `Desk Booking`
- **Book for me only** (Yes/No): when Yes, name and email are locked to the signed-in user; users cannot book for someone else
- **Allow booking any day** (Yes/No): when Yes, users can pick a future weekday; when No, booking is for today only

## Business rules

- Bookings are allowed on weekdays only
- When future booking is disabled, bookings are for today only
- Start and end times must fall between 7:00 AM and 6:00 PM
- End time must be after start time
- A desk cannot be booked twice for overlapping times on the same date
- Users provide name and email when booking (no SharePoint account required)
- Users can look up and cancel bookings by entering the same email address
- Cancelled bookings remain in the list for audit history

## SPFx version

![version](https://img.shields.io/badge/version-1.22.2-green.svg)
