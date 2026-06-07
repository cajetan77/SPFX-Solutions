import { BookingStatus, IBookerDetails, IBooking, IBookingRequest } from '../models/IModels';
import { isPastDate, isSameDate, isToday, isWeekday, isWithinBusinessHours, timesOverlap } from './dateTimeUtils';

export interface IValidationResult {
  isValid: boolean;
  message?: string;
}

const EMAIL_PATTERN = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

export function normalizeEmail(email: string): string {
  return email.trim().toLowerCase();
}

export function parseAdminEmails(adminEmailsConfig: string): string[] {
  return adminEmailsConfig
    .split(/[,;]/)
    .map(email => normalizeEmail(email))
    .filter(email => email.length > 0);
}

export function isBookingAdmin(currentUserEmail: string, adminEmails: string[]): boolean {
  if (!currentUserEmail || adminEmails.length === 0) {
    return false;
  }

  return adminEmails.indexOf(normalizeEmail(currentUserEmail)) !== -1;
}

export function isValidCancelCode(code: string, expectedCode: string): boolean {
  const entered = code.trim();
  const expected = expectedCode.trim();
  return expected.length > 0 && entered.length > 0 && entered === expected;
}

export function validateBookerDetails(details: IBookerDetails): IValidationResult {
  const name = details.name.trim();
  const email = details.email.trim();

  if (!name) {
    return { isValid: false, message: 'Please enter your name.' };
  }

  if (!email) {
    return { isValid: false, message: 'Please enter your email address.' };
  }

  if (!EMAIL_PATTERN.test(email)) {
    return { isValid: false, message: 'Please enter a valid email address.' };
  }

  return { isValid: true };
}

export function validateBookingRequest(
  request: IBookingRequest,
  options?: { todayOnly?: boolean; requirePerson?: boolean }
): IValidationResult {
  if (options?.requirePerson && !request.bookerPersonId) {
    return { isValid: false, message: 'Please select a person to book for.' };
  }

  const bookerValidation = validateBookerDetails({
    name: request.bookerName,
    email: request.bookerEmail
  });

  if (!bookerValidation.isValid) {
    return bookerValidation;
  }

  if (options?.todayOnly && !isToday(request.bookingDate)) {
    return {
      isValid: false,
      message: 'Bookings can only be made for today.'
    };
  }

  if (!options?.todayOnly && isPastDate(request.bookingDate)) {
    return {
      isValid: false,
      message: 'Bookings cannot be made for past dates.'
    };
  }

  if (!isWeekday(request.bookingDate)) {
    return {
      isValid: false,
      message: 'Bookings are only allowed Monday through Friday.'
    };
  }

  if (!isWithinBusinessHours(request.startTime, request.endTime)) {
    return {
      isValid: false,
      message: 'Booking time must be between 7:00 AM and 6:00 PM, with the end time after the start time.'
    };
  }

  return { isValid: true };
}

export function hasConflictingBooking(
  request: IBookingRequest,
  existingBookings: IBooking[]
): boolean {
  return existingBookings.some(booking =>
    booking.bookingStatus === BookingStatus.Booked
    && booking.deskId === request.deskId
    && isSameDate(booking.bookingDate, request.bookingDate)
    && timesOverlap(request.startTime, request.endTime, booking.startTime, booking.endTime)
  );
}

export function hasBookerBookingOnDate(
  bookerEmail: string,
  bookingDate: Date,
  existingBookings: IBooking[]
): boolean {
  const normalizedEmail = normalizeEmail(bookerEmail);

  return existingBookings.some(booking =>
    booking.bookingStatus === BookingStatus.Booked
    && normalizeEmail(booking.bookerEmail) === normalizedEmail
    && isSameDate(booking.bookingDate, bookingDate)
  );
}

export function canCancelBooking(
  booking: IBooking,
  currentUserEmail: string,
  isAdmin: boolean
): boolean {
  if (booking.bookingStatus !== BookingStatus.Booked) {
    return false;
  }

  if (isAdmin) {
    return true;
  }

  const normalizedCurrent = normalizeEmail(currentUserEmail);
  return normalizeEmail(booking.bookerEmail) === normalizedCurrent
    || (!!booking.createdByEmail && normalizeEmail(booking.createdByEmail) === normalizedCurrent);
}
