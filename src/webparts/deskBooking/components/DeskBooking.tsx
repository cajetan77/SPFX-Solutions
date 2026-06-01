import * as React from 'react';
import { useCallback, useEffect, useMemo, useState } from 'react';
import {
  DatePicker,
  DayOfWeek,
  DefaultButton,
  Dropdown,
  IDropdownOption,
  Label,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  Spinner,
  SpinnerSize,
  Stack,
  Text,
  TextField
} from '@fluentui/react';
import styles from './DeskBooking.module.scss';
import type { IDeskBookingProps } from './IDeskBookingProps';
import { DeskBookingService } from '../services/DeskBookingService';
import {
  BookingStatus,
  DeskDisplayStatus,
  IBooking,
  IDeskWithStatus
} from '../models/IModels';
import {
  formatDisplayDate,
  generateTimeSlots,
  isWeekday,
  timesOverlap
} from '../utils/dateTimeUtils';
import {
  canCancelBooking,
  hasConflictingBooking,
  validateBookerDetails,
  validateBookingRequest
} from '../utils/bookingRules';

const DeskBooking: React.FC<IDeskBookingProps> = (props) => {
  const { context, deskMasterListTitle, deskBookingListTitle } = props;

  const service = useMemo(
    () => new DeskBookingService(context, deskMasterListTitle, deskBookingListTitle),
    [context, deskMasterListTitle, deskBookingListTitle]
  );

  const timeSlots = useMemo(() => generateTimeSlots(), []);
  const timeOptions: IDropdownOption[] = useMemo(
    () => timeSlots.map(slot => ({ key: slot, text: slot })),
    [timeSlots]
  );

  const [bookerName, setBookerName] = useState<string>('');
  const [bookerEmail, setBookerEmail] = useState<string>('');
  const [selectedDate, setSelectedDate] = useState<Date>(() => getNextWeekday(new Date()));
  const [startTime, setStartTime] = useState<string>('09:00');
  const [endTime, setEndTime] = useState<string>('17:00');
  const [desks, setDesks] = useState<IDeskWithStatus[]>([]);
  const [yourBookings, setYourBookings] = useState<IBooking[]>([]);
  const [bookingsLoaded, setBookingsLoaded] = useState<boolean>(false);
  const [loading, setLoading] = useState<boolean>(true);
  const [loadingBookings, setLoadingBookings] = useState<boolean>(false);
  const [actionInProgress, setActionInProgress] = useState<boolean>(false);
  const [errorMessage, setErrorMessage] = useState<string | undefined>(undefined);
  const [successMessage, setSuccessMessage] = useState<string | undefined>(undefined);

  const loadDesks = useCallback(async (): Promise<void> => {
    setLoading(true);
    setErrorMessage(undefined);

    try {
      const [activeDesks, bookingsForDate] = await Promise.all([
        service.getActiveDesks(),
        service.getBookingsForDate(selectedDate)
      ]);

      const activeBookings = bookingsForDate.filter(booking => booking.bookingStatus === BookingStatus.Booked);

      const desksWithStatus: IDeskWithStatus[] = activeDesks.map(desk => {
        const deskBookings = activeBookings.filter(booking => booking.deskId === desk.id);
        const isBooked = deskBookings.some(booking =>
          timesOverlap(startTime, endTime, booking.startTime, booking.endTime)
        );

        return {
          ...desk,
          displayStatus: isBooked ? DeskDisplayStatus.Booked : DeskDisplayStatus.Available
        };
      });

      setDesks(desksWithStatus);
    } catch (error) {
      setErrorMessage(getErrorMessage(error));
    } finally {
      setLoading(false);
    }
  }, [endTime, selectedDate, service, startTime]);

  const loadYourBookings = useCallback(async (): Promise<void> => {
    const validation = validateBookerDetails({ name: bookerName, email: bookerEmail });
    if (!validation.isValid) {
      setErrorMessage(validation.message);
      return;
    }

    setLoadingBookings(true);
    setErrorMessage(undefined);

    try {
      const bookings = await service.getBookingsByEmail(bookerEmail);
      setYourBookings(bookings);
      setBookingsLoaded(true);
    } catch (error) {
      setErrorMessage(getErrorMessage(error));
    } finally {
      setLoadingBookings(false);
    }
  }, [bookerEmail, bookerName, service]);

  useEffect(() => {
    loadDesks().catch(() => {
      // Error state is handled in loadDesks.
    });
  }, [loadDesks]);

  const handleDateChange = (date: Date | null | undefined): void => {
    if (!date) {
      return;
    }

    if (!isWeekday(date)) {
      setErrorMessage('Please select a weekday (Monday through Friday).');
      return;
    }

    setSelectedDate(date);
    setSuccessMessage(undefined);
    setErrorMessage(undefined);
  };

  const handleBookDesk = async (deskId: number): Promise<void> => {
    const request = {
      deskId,
      bookingDate: selectedDate,
      startTime,
      endTime,
      bookerName,
      bookerEmail
    };

    const validation = validateBookingRequest(request);
    if (!validation.isValid) {
      setErrorMessage(validation.message);
      return;
    }

    setActionInProgress(true);
    setErrorMessage(undefined);
    setSuccessMessage(undefined);

    try {
      const bookingsForDate = await service.getBookingsForDate(selectedDate);

      if (hasConflictingBooking(request, bookingsForDate)) {
        setErrorMessage('This desk is already booked for the selected date and time.');
        return;
      }

      await service.createBooking(request);
      setSuccessMessage('Desk booked successfully.');
      await loadDesks();
      if (bookingsLoaded) {
        await loadYourBookings();
      }
    } catch (error) {
      setErrorMessage(getErrorMessage(error));
    } finally {
      setActionInProgress(false);
    }
  };

  const handleCancelBooking = async (booking: IBooking): Promise<void> => {
    if (!canCancelBooking(booking, bookerEmail)) {
      setErrorMessage('You can only cancel bookings that match the email address you entered.');
      return;
    }

    setActionInProgress(true);
    setErrorMessage(undefined);
    setSuccessMessage(undefined);

    try {
      await service.cancelBooking(booking.id);
      setSuccessMessage('Booking cancelled successfully.');
      await loadDesks();
      await loadYourBookings();
    } catch (error) {
      setErrorMessage(getErrorMessage(error));
    } finally {
      setActionInProgress(false);
    }
  };

  const renderStatusBadge = (status: DeskDisplayStatus | BookingStatus): React.ReactNode => {
    const className = `${styles.statusBadge} ${styles[`status${status}`]}`;
    return <span className={className}>{status}</span>;
  };

  return (
    <section className={styles.deskBooking}>
      <Stack tokens={{ childrenGap: 16 }}>
        <Stack tokens={{ childrenGap: 4 }}>
          <Text variant="xLarge" block>Desk Booking</Text>
          <Text variant="medium" block>
            Enter your name and email to book a desk for a weekday between 7:00 AM and 6:00 PM.
          </Text>
        </Stack>

        {errorMessage && (
          <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setErrorMessage(undefined)}>
            {errorMessage}
          </MessageBar>
        )}

        {successMessage && (
          <MessageBar messageBarType={MessageBarType.success} onDismiss={() => setSuccessMessage(undefined)}>
            {successMessage}
          </MessageBar>
        )}

        <Stack horizontal wrap tokens={{ childrenGap: 16 }} verticalAlign="end" className={styles.filters}>
          <Stack.Item grow={1} styles={{ root: { minWidth: 220 } }}>
            <TextField
              label="Your name"
              value={bookerName}
              onChange={(_, value) => setBookerName(value || '')}
              required
              placeholder="Enter your full name"
            />
          </Stack.Item>
          <Stack.Item grow={1} styles={{ root: { minWidth: 220 } }}>
            <TextField
              label="Your email"
              value={bookerEmail}
              onChange={(_, value) => {
                setBookerEmail(value || '');
                setBookingsLoaded(false);
              }}
              required
              placeholder="Enter your email address"
            />
          </Stack.Item>
        </Stack>

        <Stack horizontal wrap tokens={{ childrenGap: 16 }} verticalAlign="end" className={styles.filters}>
          <Stack.Item grow={1} styles={{ root: { minWidth: 220 } }}>
            <DatePicker
              label="Booking date"
              value={selectedDate}
              onSelectDate={handleDateChange}
              firstDayOfWeek={DayOfWeek.Monday}
              isRequired
              placeholder="Select a weekday"
              ariaLabel="Booking date"
            />
          </Stack.Item>
          <Stack.Item grow={1} styles={{ root: { minWidth: 180 } }}>
            <Dropdown
              label="Start time"
              selectedKey={startTime}
              options={timeOptions}
              onChange={(_, option) => option && setStartTime(option.key as string)}
              required
            />
          </Stack.Item>
          <Stack.Item grow={1} styles={{ root: { minWidth: 180 } }}>
            <Dropdown
              label="End time"
              selectedKey={endTime}
              options={timeOptions}
              onChange={(_, option) => option && setEndTime(option.key as string)}
              required
            />
          </Stack.Item>
          <Stack.Item>
            <DefaultButton text="Refresh" onClick={() => loadDesks()} disabled={loading || actionInProgress} />
          </Stack.Item>
        </Stack>

        <Stack tokens={{ childrenGap: 12 }}>
          <Label>Available desks for {formatDisplayDate(selectedDate)}</Label>

          {loading ? (
            <Spinner size={SpinnerSize.large} label="Loading desks..." />
          ) : desks.length === 0 ? (
            <MessageBar messageBarType={MessageBarType.info}>
              No active desks found. Add desks to the &quot;{deskMasterListTitle}&quot; list or check your list configuration.
            </MessageBar>
          ) : (
            <div className={styles.deskGrid}>
              {desks.map(desk => (
                <div key={desk.id} className={styles.deskCard}>
                  <Stack tokens={{ childrenGap: 8 }}>
                    <Text variant="large" block>{desk.deskName}</Text>
                    <Text variant="small" block>{desk.location || 'No location specified'}</Text>
                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                      {renderStatusBadge(desk.displayStatus)}
                    </Stack>
                    <PrimaryButton
                      text="Book"
                      disabled={desk.displayStatus !== DeskDisplayStatus.Available || actionInProgress}
                      onClick={() => handleBookDesk(desk.id)}
                    />
                  </Stack>
                </div>
              ))}
            </div>
          )}
        </Stack>

        <Stack tokens={{ childrenGap: 12 }}>
          <Stack horizontal wrap verticalAlign="center" tokens={{ childrenGap: 12 }}>
            <Label>Your bookings</Label>
            <DefaultButton
              text="Find my bookings"
              onClick={() => loadYourBookings()}
              disabled={loadingBookings || actionInProgress}
            />
          </Stack>
          <Text variant="small" block>
            Enter your email above, then click Find my bookings to view or cancel your reservations.
          </Text>

          {loadingBookings ? (
            <Spinner size={SpinnerSize.medium} label="Loading your bookings..." />
          ) : bookingsLoaded && yourBookings.length === 0 ? (
            <MessageBar messageBarType={MessageBarType.info}>
              No bookings found for this email address.
            </MessageBar>
          ) : bookingsLoaded ? (
            <div className={styles.bookingList}>
              {yourBookings.map(booking => (
                <div key={booking.id} className={styles.bookingRow}>
                  <Stack horizontal wrap verticalAlign="center" tokens={{ childrenGap: 12 }}>
                    <Stack.Item grow={1} styles={{ root: { minWidth: 240 } }}>
                      <Text block variant="mediumPlus">{booking.deskName}</Text>
                      <Text block variant="small">
                        {formatDisplayDate(booking.bookingDate)} | {booking.startTime} - {booking.endTime}
                      </Text>
                      <Text block variant="small">{booking.bookerName} ({booking.bookerEmail})</Text>
                    </Stack.Item>
                    <Stack.Item>
                      {renderStatusBadge(booking.bookingStatus)}
                    </Stack.Item>
                    <Stack.Item>
                      {canCancelBooking(booking, bookerEmail) && (
                        <DefaultButton
                          text="Cancel"
                          disabled={actionInProgress}
                          onClick={() => handleCancelBooking(booking)}
                        />
                      )}
                    </Stack.Item>
                  </Stack>
                </div>
              ))}
            </div>
          ) : null}
        </Stack>
      </Stack>
    </section>
  );
};

function getNextWeekday(date: Date): Date {
  const next = new Date(date.getFullYear(), date.getMonth(), date.getDate());

  while (!isWeekday(next)) {
    next.setDate(next.getDate() + 1);
  }

  return next;
}

function getErrorMessage(error: unknown): string {
  if (error instanceof Error) {
    return error.message;
  }

  return 'An unexpected error occurred. Please try again.';
}

export default DeskBooking;
