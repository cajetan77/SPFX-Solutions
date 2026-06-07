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

interface ISharePointListInfo {
  Id: string;
  Title: string;
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
  BookerPerson?: {
    Id: number;
    Title?: string;
    EMail?: string;
  };
  Author?: {
    Id: number;
    Title?: string;
    EMail?: string;
  };
  Desk?: {
    Id: number;
    Title: string;
  };
}

export class DeskBookingService {
  private readonly _context: WebPartContext;
  private readonly _deskMasterListTitle: string;
  private readonly _deskBookingListTitle: string;
  private readonly _settingsListTitle: string;

  public constructor(
    context: WebPartContext,
    deskMasterListTitle: string,
    deskBookingListTitle: string,
    settingsListTitle: string
  ) {
    this._context = context;
    this._deskMasterListTitle = deskMasterListTitle;
    this._deskBookingListTitle = deskBookingListTitle;
    this._settingsListTitle = settingsListTitle;
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
    const url = `${this._context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listTitle}')/items?$select=Id,BookingDate,StartTime,EndTime,BookingStatus,BookerName,BookerEmail,BookerPerson/Id,BookerPerson/Title,BookerPerson/EMail,Author/Id,Author/Title,Author/EMail,Desk/Id,Desk/Title&$expand=Desk,BookerPerson,Author&$filter=BookingDate ge datetime'${dateFilter}T00:00:00' and BookingDate lt datetime'${nextDayFilter}T00:00:00'`;

    const response = await this._context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    await this._ensureSuccess(response, `Unable to load bookings from "${this._deskBookingListTitle}".`);

    const data = await response.json() as ISharePointListItemResponse;
    return data.value.map(item => this._mapBookingItem(item));
  }

  public async getBookingsByEmail(email: string): Promise<IBooking[]> {
    const normalizedEmail = normalizeEmail(email);
    const listTitle = encodeURIComponent(this._deskBookingListTitle);
    const url = `${this._context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listTitle}')/items?$select=Id,BookingDate,StartTime,EndTime,BookingStatus,BookerName,BookerEmail,BookerPerson/Id,BookerPerson/Title,BookerPerson/EMail,Author/Id,Author/Title,Author/EMail,Desk/Id,Desk/Title&$expand=Desk,BookerPerson,Author&$filter=(BookerEmail eq '${normalizedEmail}' or Author/EMail eq '${normalizedEmail}')&$orderby=BookingDate desc,StartTime asc`;

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
      BookerPersonId: request.bookerPersonId || null,
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

  public async resolveUserId(loginName: string): Promise<number> {
    const url = `${this._context.pageContext.web.absoluteUrl}/_api/web/ensureuser`;

    const response = await this._context.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata',
          'odata-version': ''
        },
        body: JSON.stringify({ logonName: loginName })
      }
    );

    await this._ensureSuccess(response, 'Unable to resolve the selected user.');
    const data = await response.json() as { Id: number };
    return data.Id;
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

  public async getCancelConfirmationCode(): Promise<string> {
    const listBaseUrl = await this._getListBaseUrl(this._settingsListTitle);
    const url = `${listBaseUrl}/items?$top=1&$orderby=Id asc`;

    const response = await this._context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    await this._ensureSuccess(
      response,
      `Unable to load settings from "${this._settingsListTitle}". Check the list name in the web part properties and that you have read access.`
    );

    const data = await response.json() as { value: Array<Record<string, unknown>> };
    const settingsItem = data.value[0];

    if (!settingsItem) {
      return '';
    }

    return this._extractCancelCode(settingsItem);
  }

  private async _getListBaseUrl(listTitle: string): Promise<string> {
    const url = `${this._context.pageContext.web.absoluteUrl}/_api/web/lists?$select=Id,Title&$top=500`;
    const response = await this._context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    await this._ensureSuccess(response, `Unable to find the "${listTitle}" list.`);

    const data = await response.json() as { value: ISharePointListInfo[] };
    const normalizedTitle = listTitle.trim().toLowerCase();
    const match = data.value.find(list => list.Title.trim().toLowerCase() === normalizedTitle);

    if (!match) {
      throw new Error(`List "${listTitle}" was not found on this site.`);
    }

    return `${this._context.pageContext.web.absoluteUrl}/_api/web/lists('${match.Id}')`;
  }

  private _extractCancelCode(item: Record<string, unknown>): string {
    const knownFields = ['CancelCode', 'CaqncelCode', 'Caqncel_x0020_Code'];

    for (const field of knownFields) {
      const value = item[field];
      if (value !== undefined && value !== null) {
        const code = String(value).trim();
        if (code.length > 0) {
          return code;
        }
      }
    }

    for (const key of Object.keys(item)) {
      if (/cancel|caqncel/i.test(key)) {
        const value = item[key];
        if (value !== undefined && value !== null) {
          const code = String(value).trim();
          if (code.length > 0) {
            return code;
          }
        }
      }
    }

    return '';
  }

  private _mapBookingItem(item: ISharePointListItem): IBooking {
    return {
      id: item.Id,
      deskId: item.Desk?.Id || 0,
      deskName: item.Desk?.Title || 'Unknown desk',
      bookingDate: item.BookingDate ? parseSharePointDate(item.BookingDate) : new Date(),
      startTime: item.StartTime || '',
      endTime: item.EndTime || '',
      bookerName: item.BookerPerson?.Title || item.BookerName || '',
      bookerEmail: item.BookerPerson?.EMail || item.BookerEmail || '',
      createdByEmail: item.Author?.EMail || '',
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
