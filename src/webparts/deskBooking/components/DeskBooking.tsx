import * as React from 'react';
import { useCallback, useEffect, useMemo, useState } from 'react';
import {
  DatePicker,
  DayOfWeek,
  DefaultButton,
  Dialog,
  DialogFooter,
  DialogType,
  Dropdown,
  IDropdownOption,
  IPersonaProps,
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
import {
  PeoplePicker,
  PrincipalType,
  type IPeoplePickerContext
} from '@pnp/spfx-controls-react/lib/PeoplePicker';
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
  FIXED_BOOKING_END_TIME,
  FIXED_BOOKING_START_TIME,
  formatDisplayDate,
  generateTimeSlots,
  getTodayDate,
  isWeekday,
  timesOverlap
} from '../utils/dateTimeUtils';
import {
  canCancelBooking,
  hasBookerBookingOnDate,
  hasConflictingBooking,
  isBookingAdmin,
  isValidCancelCode,
  parseAdminEmails,
  validateBookerDetails,
  validateBookingRequest
} from '../utils/bookingRules';

const DeskBooking: React.FC<IDeskBookingProps> = (props) => {
  const {
    context,
    deskMasterListTitle,
    deskBookingListTitle,
    settingsListTitle,
    adminEmails,
    bookForMeOnly,
    allowAnyDayBooking
  } = props;
  const currentUserName = context.pageContext.user.displayName || '';
  const currentUserEmail = context.pageContext.user.email || '';

  const service = useMemo(
    () => new DeskBookingService(context, deskMasterListTitle, deskBookingListTitle, settingsListTitle),
    [context, deskMasterListTitle, deskBookingListTitle, settingsListTitle]
  );

  const bookingAdminEmails = useMemo(() => parseAdminEmails(adminEmails), [adminEmails]);
  const isCurrentUserAdmin = useMemo(
    () => isBookingAdmin(currentUserEmail, bookingAdminEmails),
    [bookingAdminEmails, currentUserEmail]
  );

  const timeSlots = useMemo(() => generateTimeSlots(), []);
  const timeOptions: IDropdownOption[] = useMemo(
    () => timeSlots.map(slot => ({ key: slot, text: slot })),
    [timeSlots]
  );

  const [bookerName, setBookerName] = useState<string>('');
  const [bookerEmail, setBookerEmail] = useState<string>('');
  const [bookerPersonId, setBookerPersonId] = useState<number | undefined>(undefined);

  const peoplePickerContext = useMemo(
    () => ({
      absoluteUrl: context.pageContext.web.absoluteUrl,
      msGraphClientFactory: context.msGraphClientFactory,
      spHttpClient: context.spHttpClient
    } as unknown as IPeoplePickerContext),
    [context]
  );

  const defaultSelectedUsers = useMemo(
    () => (bookForMeOnly && currentUserEmail ? [currentUserEmail] : undefined),
    [bookForMeOnly, currentUserEmail]
  );
  const [selectedDate, setSelectedDate] = useState<Date>(() =>
    allowAnyDayBooking ? getNextWeekday(new Date()) : getTodayDate()
  );
  const isBookingDay = isWeekday(allowAnyDayBooking ? selectedDate : getTodayDate());
  const [startTime, setStartTime] = useState<string>(FIXED_BOOKING_START_TIME);
  const [endTime, setEndTime] = useState<string>(FIXED_BOOKING_END_TIME);

  const effectiveStartTime = allowAnyDayBooking ? startTime : FIXED_BOOKING_START_TIME;
  const effectiveEndTime = allowAnyDayBooking ? endTime : FIXED_BOOKING_END_TIME;
  const [desks, setDesks] = useState<IDeskWithStatus[]>([]);
  const [bookingsForSelectedDate, setBookingsForSelectedDate] = useState<IBooking[]>([]);
  const [yourBookings, setYourBookings] = useState<IBooking[]>([]);
  const [bookingsLoaded, setBookingsLoaded] = useState<boolean>(false);
  const [loading, setLoading] = useState<boolean>(true);
  const [loadingBookings, setLoadingBookings] = useState<boolean>(false);
  const [actionInProgress, setActionInProgress] = useState<boolean>(false);
  const [errorMessage, setErrorMessage] = useState<string | undefined>(undefined);
  const [successMessage, setSuccessMessage] = useState<string | undefined>(undefined);
  const [cancelDialogVisible, setCancelDialogVisible] = useState<boolean>(false);
  const [cancelCodeInput, setCancelCodeInput] = useState<string>('');
  const [cancelCodeError, setCancelCodeError] = useState<string | undefined>(undefined);
  const [expectedCancelCode, setExpectedCancelCode] = useState<string>('');
  const [pendingCancelBooking, setPendingCancelBooking] = useState<IBooking | undefined>(undefined);

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
        const overlappingBooking = deskBookings.find(booking =>
          timesOverlap(effectiveStartTime, effectiveEndTime, booking.startTime, booking.endTime)
        );

        return {
          ...desk,
          displayStatus: overlappingBooking ? DeskDisplayStatus.Booked : DeskDisplayStatus.Available,
          bookedByName: overlappingBooking?.bookerName,
          activeBookingId: overlappingBooking?.id
        };
      });

      setDesks(desksWithStatus);
      setBookingsForSelectedDate(bookingsForDate);
    } catch (error) {
      setErrorMessage(getErrorMessage(error));
    } finally {
      setLoading(false);
    }
  }, [effectiveEndTime, effectiveStartTime, selectedDate, service]);

  const bookerAlreadyBookedOnDate = useMemo(
    () => !!bookerEmail && hasBookerBookingOnDate(bookerEmail, selectedDate, bookingsForSelectedDate),
    [bookerEmail, bookingsForSelectedDate, selectedDate]
  );

  const handleBookerPersonChange = useCallback(async (users: IPersonaProps[]): Promise<void> => {
    if (!users?.length) {
      setBookerPersonId(undefined);
      setBookerName('');
      setBookerEmail('');
      setBookingsLoaded(false);
      return;
    }

    const user = users[0];
    setBookerName(user.text || '');
    setBookerEmail(user.secondaryText || '');
    setBookingsLoaded(false);
    setErrorMessage(undefined);

    const numericId = Number(user.id);
    if (!isNaN(numericId) && numericId > 0) {
      setBookerPersonId(numericId);
      return;
    }

    const pickerUser = user as IPersonaProps & { loginName?: string };
    const loginName = pickerUser.loginName || user.id;
    if (!loginName) {
      setBookerPersonId(undefined);
      return;
    }

    try {
      const id = await service.resolveUserId(String(loginName));
      setBookerPersonId(id);
    } catch (error) {
      setBookerPersonId(undefined);
      setErrorMessage(getErrorMessage(error));
    }
  }, [service]);

  const loadYourBookings = useCallback(async (): Promise<void> => {
    const validation = validateBookerDetails({ name: bookerName, email: bookerEmail });

    if (!bookerPersonId) {
      setErrorMessage('Please select a person.');
      return;
    }

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
  }, [bookerEmail, bookerName, bookerPersonId, service]);

  useEffect(() => {
    if (!bookForMeOnly) {
      setBookerName('');
      setBookerEmail('');
      setBookerPersonId(undefined);
      setBookingsLoaded(false);
      return;
    }

    setBookerName(currentUserName);
    setBookerEmail(currentUserEmail);
    setBookingsLoaded(false);

    const loginName = context.pageContext.user.loginName;
    if (!loginName) {
      setBookerPersonId(undefined);
      return;
    }

    let cancelled = false;

    service.resolveUserId(loginName)
      .then(id => {
        if (!cancelled) {
          setBookerPersonId(id);
        }
      })
      .catch(() => {
        if (!cancelled) {
          setBookerPersonId(undefined);
        }
      });

    return () => {
      cancelled = true;
    };
  }, [bookForMeOnly, context.pageContext.user.loginName, currentUserEmail, currentUserName, service]);

  useEffect(() => {
    if (!allowAnyDayBooking) {
      setSelectedDate(getTodayDate());
      setStartTime(FIXED_BOOKING_START_TIME);
      setEndTime(FIXED_BOOKING_END_TIME);
    }
  }, [allowAnyDayBooking]);

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
      startTime: effectiveStartTime,
      endTime: effectiveEndTime,
      bookerName,
      bookerEmail,
      bookerPersonId
    };

    if (!allowAnyDayBooking && !isBookingDay) {
      setErrorMessage('Desk booking is only available on weekdays (Monday through Friday).');
      return;
    }

    const validation = validateBookingRequest(request, {
      todayOnly: !allowAnyDayBooking,
      requirePerson: true
    });
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

      if (hasBookerBookingOnDate(bookerEmail, selectedDate, bookingsForDate)) {
        setErrorMessage('This person already has a desk booked for the selected date. Cancel that booking before booking another desk.');
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

  const openCancelDialog = async (booking: IBooking): Promise<void> => {
    if (!canCancelBooking(booking, currentUserEmail, isCurrentUserAdmin)) {
      setErrorMessage('You do not have permission to cancel this booking.');
      return;
    }

    setPendingCancelBooking(booking);
    setCancelCodeInput('');
    setCancelCodeError(undefined);
    setExpectedCancelCode('');
    setCancelDialogVisible(true);

    try {
      const code = await service.getCancelConfirmationCode();
      setExpectedCancelCode(code);

      if (!code) {
        setCancelCodeError('Cancellation code is not configured in the settings list.');
      }
    } catch (error) {
      setCancelCodeError(getErrorMessage(error));
    }
  };

  const closeCancelDialog = (): void => {
    setCancelDialogVisible(false);
    setPendingCancelBooking(undefined);
    setCancelCodeInput('');
    setCancelCodeError(undefined);
    setExpectedCancelCode('');
  };

  const handleConfirmCancel = async (): Promise<void> => {
    if (!pendingCancelBooking) {
      return;
    }

    if (!isValidCancelCode(cancelCodeInput, expectedCancelCode)) {
      setCancelCodeError('Invalid cancellation code. Please try again.');
      return;
    }

    setActionInProgress(true);
    setErrorMessage(undefined);
    setSuccessMessage(undefined);

    try {
      await service.cancelBooking(pendingCancelBooking.id);
      closeCancelDialog();
      setSuccessMessage('Booking cancelled successfully.');
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
            {bookForMeOnly
              ? `Book a desk for yourself${allowAnyDayBooking ? ' on a weekday between 7:00 AM and 6:00 PM' : ' for today only (9:00 AM – 5:00 PM)'}.`
              : `Select a person to book a desk for${allowAnyDayBooking ? ' on any future weekday between 7:00 AM and 6:00 PM' : ' for today only (9:00 AM – 5:00 PM)'}.`}
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
          <Stack.Item grow={1} styles={{ root: { minWidth: 320, maxWidth: 480 } }}>
            <PeoplePicker
              key={bookForMeOnly ? 'booker-picker-me' : 'booker-picker-any'}
              context={peoplePickerContext}
              ensureUser={true}
              titleText={bookForMeOnly ? 'Booked for' : 'Book for'}
              personSelectionLimit={1}
              showtooltip={true}
              required={true}
              disabled={bookForMeOnly || actionInProgress}
              defaultSelectedUsers={defaultSelectedUsers}
              principalTypes={[PrincipalType.User]}
              resolveDelay={300}
              onChange={(users) => {
                if (bookForMeOnly) {
                  return;
                }

                handleBookerPersonChange(users).catch(() => {
                  // Error state is handled in handleBookerPersonChange.
                });
              }}
            />
          </Stack.Item>
        </Stack>

        <Stack horizontal wrap tokens={{ childrenGap: 16 }} verticalAlign="end" className={styles.filters}>
          <Stack.Item grow={1} styles={{ root: { minWidth: 220 } }}>
            {allowAnyDayBooking ? (
              <DatePicker
                label="Booking date"
                value={selectedDate}
                onSelectDate={handleDateChange}
                firstDayOfWeek={DayOfWeek.Monday}
                minDate={getTodayDate()}
                isRequired
                placeholder="Select a weekday"
                ariaLabel="Booking date"
              />
            ) : (
              <TextField
                label="Booking date"
                value={formatDisplayDate(getTodayDate())}
                readOnly
              />
            )}
          </Stack.Item>
          <Stack.Item grow={1} styles={{ root: { minWidth: 180 } }}>
            {allowAnyDayBooking ? (
              <Dropdown
                label="Start time"
                selectedKey={startTime}
                options={timeOptions}
                onChange={(_, option) => option && setStartTime(option.key as string)}
                required
              />
            ) : (
              <TextField
                label="Start time"
                value={FIXED_BOOKING_START_TIME}
                readOnly
              />
            )}
          </Stack.Item>
          <Stack.Item grow={1} styles={{ root: { minWidth: 180 } }}>
            {allowAnyDayBooking ? (
              <Dropdown
                label="End time"
                selectedKey={endTime}
                options={timeOptions}
                onChange={(_, option) => option && setEndTime(option.key as string)}
                required
              />
            ) : (
              <TextField
                label="End time"
                value={FIXED_BOOKING_END_TIME}
                readOnly
              />
            )}
          </Stack.Item>
          <Stack.Item>
            <DefaultButton text="Refresh" onClick={() => loadDesks()} disabled={loading || actionInProgress} />
          </Stack.Item>
        </Stack>

        <Stack tokens={{ childrenGap: 12 }}>
          <Label>Available desks for {formatDisplayDate(selectedDate)}</Label>

          {!allowAnyDayBooking && !isBookingDay ? (
            <MessageBar messageBarType={MessageBarType.warning}>
              Desk booking is only available on weekdays. Today is not a booking day.
            </MessageBar>
          ) : bookerAlreadyBookedOnDate ? (
            <MessageBar messageBarType={MessageBarType.warning}>
              {bookerName || 'This person'} already has a desk booked for {formatDisplayDate(selectedDate)}. Cancel that booking before booking another desk.
            </MessageBar>
          ) : loading ? (
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
                    {desk.displayStatus === DeskDisplayStatus.Booked && desk.bookedByName && (
                      <Text variant="small" block className={styles.bookedBy}>
                        Booked by {desk.bookedByName}
                      </Text>
                    )}
                    <Stack horizontal tokens={{ childrenGap: 8 }} className={styles.cardActions}>
                      <PrimaryButton
                        text="Book"
                        disabled={
                          (!allowAnyDayBooking && !isBookingDay)
                          || !bookerPersonId
                          || bookerAlreadyBookedOnDate
                          || desk.displayStatus !== DeskDisplayStatus.Available
                          || actionInProgress
                        }
                        onClick={() => handleBookDesk(desk.id)}
                      />
                      {isCurrentUserAdmin && (
                        <DefaultButton
                          text="Cancel booking"
                          disabled={
                            desk.displayStatus !== DeskDisplayStatus.Booked
                            || !desk.activeBookingId
                            || actionInProgress
                          }
                          onClick={() => {
                            openCancelDialog({
                              id: desk.activeBookingId!,
                              deskId: desk.id,
                              deskName: desk.deskName,
                              bookingDate: selectedDate,
                              startTime: effectiveStartTime,
                              endTime: effectiveEndTime,
                              bookerName: desk.bookedByName || '',
                              bookerEmail: '',
                              bookingStatus: BookingStatus.Booked
                            }).catch(() => {
                              // Error state is handled in openCancelDialog.
                            });
                          }}
                        />
                      )}
                    </Stack>
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
            {isCurrentUserAdmin
              ? 'As a booking admin, view bookings here and cancel from the desk cards above.'
              : bookForMeOnly
                ? 'Your account is shown above. Click Find my bookings to view your reservations.'
                : 'Select a person above, then click Find my bookings to view their reservations.'}
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
                  </Stack>
                </div>
              ))}
            </div>
          ) : null}
        </Stack>
      </Stack>

      <Dialog
        hidden={!cancelDialogVisible}
        onDismiss={closeCancelDialog}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Confirm cancellation',
          subText: pendingCancelBooking
            ? `Enter the cancellation code to cancel the booking for ${pendingCancelBooking.deskName}.`
            : 'Enter the cancellation code to cancel this booking.'
        }}
        modalProps={{ isBlocking: true }}
      >
        <TextField
          label="Cancellation code"
          value={cancelCodeInput}
          onChange={(_, value) => {
            setCancelCodeInput(value || '');
            setCancelCodeError(undefined);
          }}
          required
          disabled={actionInProgress}
          errorMessage={cancelCodeError}
        />
        <DialogFooter>
          <PrimaryButton
            text="Confirm cancel"
            onClick={() => {
              handleConfirmCancel().catch(() => {
                // Error state is handled in handleConfirmCancel.
              });
            }}
            disabled={actionInProgress || !cancelCodeInput.trim()}
          />
          <DefaultButton text="Close" onClick={closeCancelDialog} disabled={actionInProgress} />
        </DialogFooter>
      </Dialog>
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
