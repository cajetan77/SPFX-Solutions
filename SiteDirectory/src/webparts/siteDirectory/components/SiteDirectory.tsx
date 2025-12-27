import * as React from 'react';
import styles from './SiteDirectory.module.scss';
import type { ISiteDirectoryProps } from './ISiteDirectoryProps';
import { escape } from '@microsoft/sp-lodash-subset';
import HubSiteTree from './HubSiteTree';
import { IHubSite } from './IHubSite';
import { HubSiteService } from '../services/HubSiteService';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';

interface ISiteDirectoryState {
  hubSites: IHubSite[];
  loading: boolean;
  error: string | undefined;
}

export default class SiteDirectory extends React.Component<ISiteDirectoryProps, ISiteDirectoryState> {
  private _hubSiteService: HubSiteService;

  constructor(props: ISiteDirectoryProps) {
    super(props);
    this.state = {
      hubSites: [],
      loading: true,
      error: undefined
    };
    this._hubSiteService = new HubSiteService(props.spHttpClient, props.currentWebUrl);
  }

  public componentDidMount(): void {
    void this._loadHubSites();
  }

  private async _loadHubSites(): Promise<void> {
    try {
      this.setState({ loading: true, error: undefined });
      const hubSites = await this._hubSiteService.getHubSites();
      this.setState({ hubSites, loading: false });
    } catch (error) {
      console.error('Error loading hub sites:', error);
      this.setState({
        loading: false,
        error: error instanceof Error ? error.message : 'An error occurred while loading hub sites.'
      });
    }
  }

  public render(): React.ReactElement<ISiteDirectoryProps> {
    const {
      description,
      hasTeamsContext,
      numberOfColumns
    } = this.props;

    const { hubSites, loading, error } = this.state;

    return (
      <section className={`${styles.siteDirectory} ${hasTeamsContext ? styles.teams : ''}`}>
        {description && (
          <div className={styles.description}>
            {escape(description)}
          </div>
        )}
        <div className={styles.treeContainer}>
          {loading && (
            <div className={styles.loadingContainer}>
              <Spinner size={SpinnerSize.medium} label="Loading hub sites..." />
            </div>
          )}
          {error && (
            <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
              {error}
            </MessageBar>
          )}
          {!loading && !error && hubSites.length === 0 && (
            <MessageBar messageBarType={MessageBarType.info}>
              No hub sites found.
            </MessageBar>
          )}
          {!loading && !error && hubSites.length > 0 && (
            <HubSiteTree hubSites={hubSites} numberOfColumns={numberOfColumns} />
          )}
        </div>
      </section>
    );
  }
}
