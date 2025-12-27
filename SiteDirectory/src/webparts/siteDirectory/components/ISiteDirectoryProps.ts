import { SPHttpClient } from '@microsoft/sp-http';

export interface ISiteDirectoryProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  spHttpClient: SPHttpClient;
  currentWebUrl: string;
  numberOfColumns: number;
}
