import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IHubSite, IAssociatedSite } from '../components/IHubSite';

export interface IHubSiteService {
    getHubSites(): Promise<IHubSite[]>;
}

export class HubSiteService implements IHubSiteService {
    private _spHttpClient: SPHttpClient;
    private _currentWebUrl: string;

    constructor(spHttpClient: SPHttpClient, currentWebUrl: string) {
        this._spHttpClient = spHttpClient;
        this._currentWebUrl = currentWebUrl;
    }

    public async getHubSites(): Promise<IHubSite[]> {
        try {
            // Get all hub sites
            const hubSitesUrl = `${this._currentWebUrl}/_api/HubSites`;
            const hubSitesResponse: SPHttpClientResponse = await this._spHttpClient.get(
                hubSitesUrl,
                SPHttpClient.configurations.v1
            );

            if (!hubSitesResponse.ok) {
                throw new Error(`Failed to fetch hub sites: ${hubSitesResponse.statusText}`);
            }

            const hubSitesData = await hubSitesResponse.json();
            const hubSites = hubSitesData.value || [];

            // Verify each site is actually a hub site and get associated sites
            const hubSitesWithAssociations: IHubSite[] = await Promise.all(
                hubSites.map(async (hubSite: any) => {
                    // Verify this is actually a hub site
                    const isHubSite = await this.verifyIsHubSite(hubSite.SiteUrl);
                    if (!isHubSite) {
                        return null;
                    }

                    const associatedSites = await this.getAssociatedSites(hubSite.SiteUrl, hubSite.ID);

                    return {
                        id: hubSite.ID || hubSite.SiteId,
                        title: hubSite.Title || 'Untitled Hub Site',
                        url: hubSite.SiteUrl || '',
                        description: hubSite.Description,
                        associatedSites: associatedSites,
                        isHighlighted: false // Can be customized based on business logic
                    };
                })
            );

            // Filter out null values (sites that weren't actually hub sites)
            return hubSitesWithAssociations.filter((hub): hub is IHubSite => hub !== null);
        } catch (error) {
            console.error('Error fetching hub sites:', error);
            throw error;
        }
    }

    private async verifyIsHubSite(siteUrl: string): Promise<boolean> {
        try {
            // Check if the site is actually a hub by checking its web properties
            // A hub site will have HubSiteId that matches its own SiteId
            const webPropertiesUrl = `${siteUrl}/_api/web?$select=Id,HubSiteId`;
            const webPropertiesResponse: SPHttpClientResponse = await this._spHttpClient.get(
                webPropertiesUrl,
                SPHttpClient.configurations.v1
            );

            if (!webPropertiesResponse.ok) {
                return false;
            }

            const webData = await webPropertiesResponse.json();
            const siteId = webData.Id;
            const hubSiteId = webData.HubSiteId;

            // A hub site has its HubSiteId equal to its own SiteId (it points to itself)
            // Or HubSiteId is null/empty which also indicates it's a hub
            return hubSiteId === siteId || !hubSiteId || hubSiteId === '00000000-0000-0000-0000-000000000000';
        } catch (error) {
            console.warn(`Error verifying hub site for ${siteUrl}:`, error);
            // If verification fails, we'll trust the HubSites endpoint
            return true;
        }
    }

    private async getAssociatedSites(hubSiteUrl: string, hubSiteId: string): Promise<IAssociatedSite[]> {
        try {
            // Use HubSiteData endpoint first to get associated sites for this hub
            let associatedSites: IAssociatedSite[] = [];

            try {
                const hubSiteDataUrl = `${hubSiteUrl}/_api/web/HubSiteData`;
                const hubSiteDataResponse: SPHttpClientResponse = await this._spHttpClient.get(
                    hubSiteDataUrl,
                    SPHttpClient.configurations.v1
                );

                if (hubSiteDataResponse.ok) {
                    const hubSiteDataResult = await hubSiteDataResponse.json();
                    const hubSiteData = JSON.parse(hubSiteDataResult.value);
                    const associatedSitesData = hubSiteData.AssociatedSites || [];

                    // Filter out the hub site itself and verify sites are actually associated
                    associatedSites = associatedSitesData
                        .filter((site: any) => {
                            const siteUrl = site.SiteUrl || '';
                            // Exclude the hub site itself
                            if (siteUrl.toLowerCase() === hubSiteUrl.toLowerCase()) {
                                return false;
                            }
                            // Only include sites that have valid URLs
                            return siteUrl.length > 0;
                        })
                        .map((site: any) => ({
                            id: site.SiteId || '',
                            title: site.Title || 'Untitled Site',
                            url: site.SiteUrl || ''
                        }));
                }
            } catch (hubSiteDataError) {
                console.warn(`HubSiteData API failed for ${hubSiteUrl}:`, hubSiteDataError);
            }

            // Also use Search API to find all sites associated with this hub
            // This provides a more comprehensive list using DepartmentId
            try {
                const searchQuery = `DepartmentId:"${hubSiteId}" AND contentclass:STS_Site AND -Path:"${hubSiteUrl}"`;
                const searchUrl = `${this._currentWebUrl}/_api/search/query?querytext='${encodeURIComponent(searchQuery)}'&selectproperties='Title,Path,SiteId'&rowlimit=500`;

                const searchResponse: SPHttpClientResponse = await this._spHttpClient.get(
                    searchUrl,
                    SPHttpClient.configurations.v1
                );

                if (searchResponse.ok) {
                    const searchData = await searchResponse.json();
                    const rows = searchData?.PrimaryQueryResult?.RelevantResults?.Table?.Rows || [];

                    const searchSites: IAssociatedSite[] = rows.map((row: any) => {
                        const cells = row.Cells || [];
                        const getCellValue = (key: string): string => {
                            const cell = cells.find((c: any) => c.Key === key);
                            return cell ? cell.Value : '';
                        };

                        return {
                            id: getCellValue('SiteId') || '',
                            title: getCellValue('Title') || 'Untitled Site',
                            url: getCellValue('Path') || ''
                        };
                    });

                    // Merge results from both APIs, removing duplicates based on URL
                    const existingUrls = new Set(associatedSites.map(s => s.url.toLowerCase()));
                    const newSites = searchSites.filter(s => {
                        const urlLower = s.url.toLowerCase();
                        if (!existingUrls.has(urlLower) && urlLower !== hubSiteUrl.toLowerCase()) {
                            existingUrls.add(urlLower);
                            return true;
                        }
                        return false;
                    });

                    associatedSites = [...associatedSites, ...newSites];
                }
            } catch (searchError) {
                console.warn(`Search API failed for hub ${hubSiteId}:`, searchError);
            }

            return associatedSites;
        } catch (error) {
            console.error(`Error fetching associated sites for hub ${hubSiteId}:`, error);
            return [];
        }
    }
}
