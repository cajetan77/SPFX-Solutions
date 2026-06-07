declare interface IDeskBookingWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DeskMasterListTitleLabel: string;
  DeskBookingListTitleLabel: string;
  BookingOptionsGroupName: string;
  BookForMeOnlyLabel: string;
  AllowAnyDayBookingLabel: string;
  AdminGroupName: string;
  SettingsListTitleLabel: string;
  AdminEmailsLabel: string;
  AdminEmailsDescription: string;
  ToggleOnText: string;
  ToggleOffText: string;
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
