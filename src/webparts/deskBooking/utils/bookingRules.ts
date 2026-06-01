import { BookingStatus, IBookerDetails, IBooking, IBookingRequest } from '../models/IModels';
import { isSameDate, isWeekday, isWithinBusinessHours, timesOverlap } from './dateTimeUtils';

export interface IValidationResult {
  isValid: boolean;
  message?: string;
}

const EMAIL_PATTERN = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

export function normalizeEmail(email: string): string {
  return email.trim().toLowerCase();
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

export function validateBookingRequest(request: IBookingRequest): IValidationResult {
  const bookerValidation = validateBookerDetails({
    name: request.bookerName,
    email: request.bookerEmail
  });

  if (!bookerValidation.isValid) {
    return bookerValidation;
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

export function canCancelBooking(
  booking: IBooking,
  bookerEmail: string
): boolean {
  return booking.bookingStatus === BookingStatus.Booked
    && normalizeEmail(booking.bookerEmail) === normalizeEmail(bookerEmail);
}
