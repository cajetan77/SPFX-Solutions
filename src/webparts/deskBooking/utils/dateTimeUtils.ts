const BUSINESS_START_MINUTES = 7 * 60;
const BUSINESS_END_MINUTES = 18 * 60;
const TIME_SLOT_INTERVAL_MINUTES = 30;

export const FIXED_BOOKING_START_TIME = '09:00';
export const FIXED_BOOKING_END_TIME = '17:00';

export function parseTimeToMinutes(time: string): number {
  const parts = time.split(':');
  if (parts.length !== 2) {
    return NaN;
  }

  const hours = parseInt(parts[0], 10);
  const minutes = parseInt(parts[1], 10);

  if (isNaN(hours) || isNaN(minutes)) {
    return NaN;
  }

  return hours * 60 + minutes;
}

function padTwoDigits(value: number): string {
  return value < 10 ? `0${value}` : `${value}`;
}

export function formatMinutesToTime(totalMinutes: number): string {
  const hours = Math.floor(totalMinutes / 60);
  const minutes = totalMinutes % 60;
  return `${padTwoDigits(hours)}:${padTwoDigits(minutes)}`;
}

export function generateTimeSlots(): string[] {
  const slots: string[] = [];

  for (let minutes = BUSINESS_START_MINUTES; minutes <= BUSINESS_END_MINUTES; minutes += TIME_SLOT_INTERVAL_MINUTES) {
    slots.push(formatMinutesToTime(minutes));
  }

  return slots;
}

export function isWeekday(date: Date): boolean {
  const day = date.getDay();
  return day >= 1 && day <= 5;
}

export function isWithinBusinessHours(startTime: string, endTime: string): boolean {
  const startMinutes = parseTimeToMinutes(startTime);
  const endMinutes = parseTimeToMinutes(endTime);

  if (isNaN(startMinutes) || isNaN(endMinutes)) {
    return false;
  }

  return startMinutes >= BUSINESS_START_MINUTES
    && endMinutes <= BUSINESS_END_MINUTES
    && startMinutes < endMinutes;
}

export function timesOverlap(
  startA: string,
  endA: string,
  startB: string,
  endB: string
): boolean {
  const startAMinutes = parseTimeToMinutes(startA);
  const endAMinutes = parseTimeToMinutes(endA);
  const startBMinutes = parseTimeToMinutes(startB);
  const endBMinutes = parseTimeToMinutes(endB);

  return startAMinutes < endBMinutes && startBMinutes < endAMinutes;
}

export function toDateOnlyIsoString(date: Date): string {
  const year = date.getFullYear();
  const month = padTwoDigits(date.getMonth() + 1);
  const day = padTwoDigits(date.getDate());
  return `${year}-${month}-${day}`;
}

export function toLocalDateOnly(date: Date): Date {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate(), 12, 0, 0, 0);
}

export function isSameDate(a: Date, b: Date): boolean {
  return toDateOnlyIsoString(a) === toDateOnlyIsoString(b);
}

export function getTodayDate(): Date {
  return toLocalDateOnly(new Date());
}

export function isToday(date: Date): boolean {
  return isSameDate(date, getTodayDate());
}

export function isPastDate(date: Date): boolean {
  return toDateOnlyIsoString(date) < toDateOnlyIsoString(getTodayDate());
}

export function toSharePointBookingDate(date: Date): string {
  return `${toDateOnlyIsoString(date)}T00:00:00`;
}

export function parseSharePointDate(value: string): Date {
  const parsed = new Date(value);

  if (isNaN(parsed.getTime())) {
    return toLocalDateOnly(new Date());
  }

  return toLocalDateOnly(parsed);
}

export function formatDisplayDate(date: Date): string {
  return date.toLocaleDateString(undefined, {
    weekday: 'short',
    year: 'numeric',
    month: 'short',
    day: 'numeric'
  });
}
