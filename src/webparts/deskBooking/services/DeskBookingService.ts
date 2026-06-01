import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
  BookingStatus,
  IBooking,
  IBookingRequest,
  IDesk
} from '../models/IModels';
import { normalizeEmail } from '../utils/bookingRules';
import { parseSharePointDate, toDateOnlyIsoString, toSharePointBookingDate } from '../utils/dateTimeUtils';

interface ISharePointListItemResponse {
  value: ISharePointListItem[];
}

interface ISharePointListItem {
  Id: number;
  Title?: string;
  Location?: string;
  IsActive?: boolean;
  BookingDate?: string;
  StartTime?: string;
  EndTime?: string;
  BookingStatus?: string;
  BookerName?: string;
  BookerEmail?: string;
  Desk?: {
    Id: number;
    Title: string;
  };
}

export class DeskBookingService {
  private readonly _context: WebPartContext;
  private readonly _deskMasterListTitle: string;
  private readonly _deskBookingListTitle: string;

  public constructor(
    context: WebPartContext,
    deskMasterListTitle: string,
    deskBookingListTitle: string
  ) {
    this._context = context;
    this._deskMasterListTitle = deskMasterListTitle;
    this._deskBookingListTitle = deskBookingListTitle;
  }

  public async getActiveDesks(): Promise<IDesk[]> {
    const listTitle = encodeURIComponent(this._deskMasterListTitle);
    const url = `${this._context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listTitle}')/items?$select=Id,Title,Location,IsActive&$filter=IsActive eq 1&$orderby=Title asc`;

    const response = await this._context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    await this._ensureSuccess(response, `Unable to load desks from "${this._deskMasterListTitle}".`);

    const data = await response.json() as ISharePointListItemResponse;
    return data.value.map(item => ({
      id: item.Id,
      deskName: item.Title || '',
      location: item.Location || '',
      isActive: !!item.IsActive
    }));
  }

  public async getBookingsForDate(bookingDate: Date): Promise<IBooking[]> {
    const dateFilter = toDateOnlyIsoString(bookingDate);
    const nextDay = new Date(bookingDate.getFullYear(), bookingDate.getMonth(), bookingDate.getDate() + 1);
    const nextDayFilter = toDateOnlyIsoString(nextDay);
    const listTitle = encodeURIComponent(this._deskBookingListTitle);
    const url = `${this._context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listTitle}')/items?$select=Id,BookingDate,StartTime,EndTime,BookingStatus,BookerName,BookerEmail,Desk/Id,Desk/Title&$expand=Desk&$filter=BookingDate ge datetime'${dateFilter}T00:00:00' and BookingDate lt datetime'${nextDayFilter}T00:00:00'`;

    const response = await this._context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    await this._ensureSuccess(response, `Unable to load bookings from "${this._deskBookingListTitle}".`);

    const data = await response.json() as ISharePointListItemResponse;
    return data.value.map(item => this._mapBookingItem(item));
  }

  public async getBookingsByEmail(email: string): Promise<IBooking[]> {
    const normalizedEmail = normalizeEmail(email);
    const listTitle = encodeURIComponent(this._deskBookingListTitle);
    const url = `${this._context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listTitle}')/items?$select=Id,BookingDate,StartTime,EndTime,BookingStatus,BookerName,BookerEmail,Desk/Id,Desk/Title&$expand=Desk&$filter=BookerEmail eq '${normalizedEmail}'&$orderby=BookingDate desc,StartTime asc`;

    const response = await this._context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    await this._ensureSuccess(response, 'Unable to load bookings for this email address.');

    const data = await response.json() as ISharePointListItemResponse;
    return data.value.map(item => this._mapBookingItem(item));
  }

  public async createBooking(request: IBookingRequest): Promise<void> {
    const listTitle = encodeURIComponent(this._deskBookingListTitle);
    const url = `${this._context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listTitle}')/items`;

    const body = {
      DeskId: request.deskId,
      BookingDate: toSharePointBookingDate(request.bookingDate),
      StartTime: request.startTime,
      EndTime: request.endTime,
      BookerName: request.bookerName.trim(),
      BookerEmail: normalizeEmail(request.bookerEmail),
      BookingStatus: BookingStatus.Booked
    };

    const response = await this._context.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata',
          'odata-version': ''
        },
        body: JSON.stringify(body)
      }
    );

    await this._ensureSuccess(response, 'Unable to create the booking.');
  }

  public async cancelBooking(bookingId: number): Promise<void> {
    const listTitle = encodeURIComponent(this._deskBookingListTitle);
    const url = `${this._context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listTitle}')/items(${bookingId})`;

    const response = await this._context.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata',
          'odata-version': '',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'MERGE'
        },
        body: JSON.stringify({
          BookingStatus: BookingStatus.Cancelled
        })
      }
    );

    await this._ensureSuccess(response, 'Unable to cancel the booking.');
  }

  private _mapBookingItem(item: ISharePointListItem): IBooking {
    return {
      id: item.Id,
      deskId: item.Desk?.Id || 0,
      deskName: item.Desk?.Title || 'Unknown desk',
      bookingDate: item.BookingDate ? parseSharePointDate(item.BookingDate) : new Date(),
      startTime: item.StartTime || '',
      endTime: item.EndTime || '',
      bookerName: item.BookerName || '',
      bookerEmail: item.BookerEmail || '',
      bookingStatus: (item.BookingStatus as BookingStatus) || BookingStatus.Booked
    };
  }

  private async _ensureSuccess(response: SPHttpClientResponse, fallbackMessage: string): Promise<void> {
    if (response.ok) {
      return;
    }

    let message = fallbackMessage;

    try {
      const errorBody = await response.json() as { error?: { message?: { value?: string } } };
      message = errorBody?.error?.message?.value || fallbackMessage;
    } catch {
      // Keep fallback message when response body is not JSON.
    }

    throw new Error(message);
  }
}
