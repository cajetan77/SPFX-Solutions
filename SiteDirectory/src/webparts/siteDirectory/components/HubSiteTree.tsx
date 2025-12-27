import * as React from 'react';
import styles from './HubSiteTree.module.scss';
import { IHubSite, IAssociatedSite } from './IHubSite';
import { Icon } from '@fluentui/react/lib/Icon';

export interface IHubSiteTreeProps {
    hubSites: IHubSite[];
    numberOfColumns: number;
}

interface IHubSiteTreeState {
    expandedHubs: Set<string>; // Track which hubs are expanded
}

export default class HubSiteTree extends React.Component<IHubSiteTreeProps, IHubSiteTreeState> {
    constructor(props: IHubSiteTreeProps) {
        super(props);
        // Initialize all hubs as collapsed
        this.state = {
            expandedHubs: new Set()
        };
    }

    private toggleHub = (hubId: string): void => {
        this.setState(prevState => {
            const newExpanded = new Set(prevState.expandedHubs);
            if (newExpanded.has(hubId)) {
                newExpanded.delete(hubId);
            } else {
                newExpanded.add(hubId);
            }
            return { expandedHubs: newExpanded };
        });
    };

    private renderAssociatedSite = (site: IAssociatedSite, index: number, total: number): React.ReactElement => {
        const isLast = index === total - 1;
        return (
            <div key={site.id} className={styles.siteItem}>

                <div className={styles.siteConnector}>
                    <div className={isLast ? styles.lineLast : styles.lineMiddle} />
                </div>
                <a href={site.url} className={styles.siteLink} target="_blank" rel="noopener noreferrer">
                    <Icon iconName="SharepointAppIcon16" className={styles.siteIcon} />
                    <span className={styles.siteTitle}>{site.title}</span>
                </a>
            </div>
        );
    };

    private renderHubSite = (hub: IHubSite, index: number): React.ReactElement => {
        const isExpanded = this.state.expandedHubs.has(hub.id);
        const hasChildren = hub.associatedSites && hub.associatedSites.length > 0;

        return (
            <div key={hub.id} className={styles.hubContainer}>
                <div className={styles.hubItem}>
                    <button
                        className={styles.expandButton}
                        onClick={() => this.toggleHub(hub.id)}
                        disabled={!hasChildren}
                        aria-expanded={isExpanded}
                        aria-label={`${isExpanded ? 'Collapse' : 'Expand'} ${hub.title}`}
                    >
                        <span className={`${styles.expandIcon} ${isExpanded ? styles.expanded : ''}`}>
                            ▶
                        </span>
                    </button>
                    <a
                        href={hub.url}
                        className={styles.hubLink}
                        target="_blank"
                        rel="noopener noreferrer"
                    >
                        <Icon iconName="SharepointAppIcon16" className={styles.hubIcon} />

                        <span className={hub.isHighlighted ? styles.hubLetterHighlighted : styles.hubLetter}>
                            {hub.title.replace(/^Hub Site\s*/i, '')}
                        </span>
                        {hasChildren && (
                            <span className={styles.siteCount}>({hub.associatedSites.length})</span>
                        )}
                    </a>
                </div>
                {isExpanded && hasChildren && (
                    <div className={styles.sitesContainer}>
                        {hub.associatedSites.map((site, siteIndex) =>
                            this.renderAssociatedSite(site, siteIndex, hub.associatedSites.length)
                        )}
                    </div>
                )}
            </div>
        );
    };

    public render(): React.ReactElement<IHubSiteTreeProps> {
        const columnStyle = { gridTemplateColumns: `repeat(${this.props.numberOfColumns}, 1fr)` };
        return (
            <div className={styles.hubSiteTree} style={columnStyle}>
                {this.props.hubSites.map((hub, index) => this.renderHubSite(hub, index))}
            </div>
        );
    }
}
