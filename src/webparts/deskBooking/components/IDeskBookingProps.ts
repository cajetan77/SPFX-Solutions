import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IDeskBookingProps {
  context: WebPartContext;
  deskMasterListTitle: string;
  deskBookingListTitle: string;
  settingsListTitle: string;
  adminEmails: string;
  bookForMeOnly: boolean;
  allowAnyDayBooking: boolean;
}
