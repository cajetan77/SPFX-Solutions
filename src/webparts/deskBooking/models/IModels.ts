export enum BookingStatus {
  Booked = 'Booked',
  Cancelled = 'Cancelled'
}

export enum DeskDisplayStatus {
  Available = 'Available',
  Booked = 'Booked',
  Cancelled = 'Cancelled'
}

export interface IDesk {
  id: number;
  deskName: string;
  location: string;
  isActive: boolean;
}

export interface IBooking {
  id: number;
  deskId: number;
  deskName: string;
  bookingDate: Date;
  startTime: string;
  endTime: string;
  bookerName: string;
  bookerEmail: string;
  createdByEmail?: string;
  bookingStatus: BookingStatus;
}

export interface IDeskWithStatus extends IDesk {
  displayStatus: DeskDisplayStatus;
  bookedByName?: string;
  activeBookingId?: number;
}

export interface IBookingRequest {
  deskId: number;
  bookingDate: Date;
  startTime: string;
  endTime: string;
  bookerName: string;
  bookerEmail: string;
  bookerPersonId?: number;
}

export interface IBookerDetails {
  name: string;
  email: string;
}
