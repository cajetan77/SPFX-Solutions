# Site Directory

## Summary

The **Site Directory** web part is a SharePoint Framework (SPFx) solution that provides a comprehensive directory view of SharePoint Hub Sites and their associated sites within your Microsoft 365 tenant. This web part helps users navigate and discover sites organized under hub sites, making it easier to find and access relevant SharePoint sites.

### Key Features

- **Hub Site Discovery**: Automatically discovers and displays all SharePoint Hub Sites in your tenant
- **Hierarchical Tree View**: Shows hub sites with their associated sites in an expandable/collapsible tree structure
- **Interactive Navigation**: Click on any hub site or associated site to navigate directly to it
- **Configurable Layout**: Choose between 1 or 2 column layouts for optimal display
- **Smart Site Detection**: Uses multiple SharePoint APIs (HubSites, HubSiteData, and Search) to ensure comprehensive site discovery
- **Loading States**: Displays loading indicators while fetching data
- **Error Handling**: Provides user-friendly error messages when issues occur
- **Theme Support**: Automatically adapts to SharePoint theme variants (light/dark mode)
- **Multi-Host Support**: Works in SharePoint pages, Teams tabs, and Teams personal apps

### How It Works

1. **Hub Site Detection**: The web part queries the SharePoint REST API to retrieve all hub sites in your tenant
2. **Site Association**: For each hub site, it fetches all associated sites using both the HubSiteData API and Search API
3. **Tree Display**: Hub sites are displayed in a grid layout (1 or 2 columns) with expand/collapse functionality
4. **Site Navigation**: Users can expand hub sites to see associated sites and click on any site to navigate to it
5. **Visual Indicators**: Shows site counts and uses icons to distinguish hub sites from associated sites

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.22.1-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

- SharePoint Hub Sites must be configured in your tenant
- Users need appropriate permissions to view hub sites and associated sites
- The web part uses SharePoint REST APIs, so users must have access to the sites they want to view

## Solution

| Solution        | Author(s) |
| --------------- | --------- |
| site-directory  |           |

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 1.1     | March 10, 2021   | Update comment  |
| 1.0     | January 29, 2021 | Initial release |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - `npm install -g @rushstack/heft`
  - `npm install`
  - `heft start`

> Include any additional steps as needed.

Other build commands can be listed using `heft --help`.

## Features

This web part demonstrates the following SharePoint Framework concepts and capabilities:

- **SharePoint REST API Integration**: Uses multiple SharePoint REST endpoints (`/_api/HubSites`, `/_api/web/HubSiteData`, `/_api/search/query`) to fetch hub site information
- **React Component Architecture**: Built with React and TypeScript for type-safe, component-based development
- **Fluent UI Components**: Leverages Microsoft Fluent UI (formerly Office UI Fabric) for consistent, accessible UI components
- **SPHttpClient Usage**: Demonstrates proper use of SPHttpClient for authenticated SharePoint API calls
- **Error Handling**: Implements comprehensive error handling with user-friendly error messages
- **Loading States**: Shows appropriate loading indicators during asynchronous operations
- **Property Pane Configuration**: Configurable web part properties (description and column layout)
- **Theme Variants Support**: Automatically adapts to SharePoint site themes
- **Multi-Host Compatibility**: Supports SharePoint pages, Teams tabs, and Teams personal apps
- **Tree View Component**: Custom tree view implementation with expand/collapse functionality

> Notice that better pictures and documentation will increase the sample usage and the value you are providing for others. Thanks for your submissions advance.

> Share your web part with others through Microsoft 365 Patterns and Practices program to get visibility and exposure. More details on the community, open-source projects and other activities from http://aka.ms/m365pnp.

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
- [Heft Documentation](https://heft.rushstack.io/)