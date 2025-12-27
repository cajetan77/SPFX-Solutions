export interface IAssociatedSite {
    id: string;
    title: string;
    url: string;
    description?: string;
}

export interface IHubSite {
    id: string;
    title: string;
    url: string;
    description?: string;
    associatedSites: IAssociatedSite[];
    isHighlighted?: boolean; // For red highlighting like in the image
}
