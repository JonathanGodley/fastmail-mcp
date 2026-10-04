import { coercePosition } from './coerce.js';
import { formatCalendarEventList } from './response-formatters.js';
import type { CalendarEventQueryResult } from './caldav-client.js';
import type { ToolContent } from './contacts-handler.js';

/**
 * The slice of the CalDAV client list_calendar_events needs, so the handler can be exercised
 * with a stub. `CalDAVCalendarClient` satisfies it structurally.
 */
export interface CalendarListClient {
  getCalendarEvents(
    calendarId?: string,
    limit?: number,
    startDate?: string,
    endDate?: string,
    position?: number,
  ): Promise<CalendarEventQueryResult>;
}

// `limit` arrives clamped: the clamp stays in the CallTool case, where tool-schema.test.ts
// checks it against the advertised cap.
export async function listCalendarEventsTool(args: any, limit: number, client: CalendarListClient): Promise<ToolContent> {
  const { calendarId, startDate, endDate } = args ?? {};
  // Rejected before the read, which is every calendar in the window.
  const position = coercePosition(args?.position);
  const result = await client.getCalendarEvents(calendarId, limit, startDate, endDate, position);
  return [{ type: 'text', text: formatCalendarEventList(result) }];
}
