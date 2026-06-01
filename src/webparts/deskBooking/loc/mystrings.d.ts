declare interface IDeskBookingWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DeskMasterListTitleLabel: string;
  DeskBookingListTitleLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;
}

declare module 'DeskBookingWebPartStrings' {
  const strings: IDeskBookingWebPartStrings;
  export = strings;
}
