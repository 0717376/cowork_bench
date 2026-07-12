import pg from 'pg';
const { Pool } = pg;

export interface CalendarEvent {
    id: string;
    summary: string | null;
    description: string | null;
    location: string | null;
    start: { dateTime: string | null; timeZone?: string };
    end: { dateTime: string | null; timeZone?: string };
    status: string | null;
    htmlLink: string | null;
    creator: any;
    organizer: any;
    attendees: any;
    recurrence: any;
    reminders: any;
    created: string | null;
    updated: string | null;
}

interface EventRow {
    id: string;
    summary: string | null;
    description: string | null;
    location: string | null;
    start_datetime: string | null;
    start_timezone: string | null;
    end_datetime: string | null;
    end_timezone: string | null;
    status: string | null;
    html_link: string | null;
    creator: any;
    organizer: any;
    attendees: any;
    recurrence: any;
    reminders: any;
    created: string | null;
    updated: string | null;
}

function formatEvent(row: EventRow): CalendarEvent {
    return {
        id: row.id,
        summary: row.summary,
        description: row.description,
        location: row.location,
        start: {
            dateTime: row.start_datetime ? new Date(row.start_datetime).toISOString() : null,
            timeZone: row.start_timezone || undefined,
        },
        end: {
            dateTime: row.end_datetime ? new Date(row.end_datetime).toISOString() : null,
            timeZone: row.end_timezone || undefined,
        },
        status: row.status,
        htmlLink: row.html_link,
        creator: row.creator,
        organizer: row.organizer,
        attendees: row.attendees,
        recurrence: row.recurrence,
        reminders: row.reminders,
        created: row.created ? new Date(row.created).toISOString() : null,
        updated: row.updated ? new Date(row.updated).toISOString() : null,
    };
}

// Offset of an IANA time zone at a given instant, in milliseconds.
function tzOffsetMs(date: Date, timeZone: string): number {
    const parts: Record<string, string> = {};
    const fmt = new Intl.DateTimeFormat('en-US', {
        timeZone,
        year: 'numeric', month: '2-digit', day: '2-digit',
        hour: '2-digit', minute: '2-digit', second: '2-digit',
        hour12: false,
    });
    for (const p of fmt.formatToParts(date)) parts[p.type] = p.value;
    const asUtc = Date.UTC(
        Number(parts.year), Number(parts.month) - 1, Number(parts.day),
        Number(parts.hour) % 24, Number(parts.minute), Number(parts.second)
    );
    return asUtc - date.getTime();
}

// Google Calendar API semantics: a naive dateTime (no Z/offset) combined with
// a timeZone is interpreted in that IANA zone and converted to UTC.
function resolveDateTime(dateTime: string | null | undefined, timeZone: string | null | undefined): string | null {
    if (!dateTime) return null;
    if (/(Z|[+-]\d{2}:?\d{2})$/i.test(dateTime) || !timeZone) return dateTime;
    try {
        const naiveAsUtc = new Date(`${dateTime}Z`);
        if (isNaN(naiveAsUtc.getTime())) return dateTime;
        let utcMs = naiveAsUtc.getTime() - tzOffsetMs(naiveAsUtc, timeZone);
        utcMs = naiveAsUtc.getTime() - tzOffsetMs(new Date(utcMs), timeZone);
        return new Date(utcMs).toISOString();
    } catch {
        return dateTime;
    }
}

export class PgCalendar {
    events: {
        insert(params: { calendarId: string; requestBody: any }): Promise<{ data: CalendarEvent }>;
        get(params: { calendarId: string; eventId: string }): Promise<{ data: CalendarEvent }>;
        patch(params: { calendarId: string; eventId: string; requestBody: any }): Promise<{ data: CalendarEvent }>;
        delete(params: { calendarId: string; eventId: string }): Promise<{ data: {} }>;
        list(params: {
            calendarId: string;
            timeMin?: string;
            timeMax?: string;
            maxResults?: number;
            orderBy?: string;
            singleEvents?: boolean;
        }): Promise<{ data: { items: CalendarEvent[] } }>;
    };

    constructor(pool: InstanceType<typeof Pool>) {
        this.events = {
            async insert({ calendarId, requestBody }) {
                const startTimeZone = requestBody.start?.timeZone || null;
                const startDateTime = resolveDateTime(requestBody.start?.dateTime, startTimeZone);
                const endTimeZone = requestBody.end?.timeZone || null;
                const endDateTime = resolveDateTime(requestBody.end?.dateTime, endTimeZone);
                const attendeesJson = requestBody.attendees
                    ? JSON.stringify(requestBody.attendees)
                    : '[]';

                const result = await pool.query(
                    `INSERT INTO gcal.events (summary, description, location, start_datetime, start_timezone, end_datetime, end_timezone, attendees)
                     VALUES ($1, $2, $3, $4, $5, $6, $7, $8::jsonb)
                     RETURNING *`,
                    [
                        requestBody.summary || null,
                        requestBody.description || null,
                        requestBody.location || null,
                        startDateTime,
                        startTimeZone,
                        endDateTime,
                        endTimeZone,
                        attendeesJson,
                    ]
                );
                return { data: formatEvent(result.rows[0]) };
            },

            async get({ calendarId, eventId }) {
                const result = await pool.query(
                    `SELECT * FROM gcal.events WHERE id = $1`,
                    [eventId]
                );
                if (result.rows.length === 0) {
                    throw new Error(`Event not found: ${eventId}`);
                }
                return { data: formatEvent(result.rows[0]) };
            },

            async patch({ calendarId, eventId, requestBody }) {
                const setClauses: string[] = [];
                const values: any[] = [];
                let paramIndex = 1;

                if (requestBody.summary !== undefined) {
                    setClauses.push(`summary = $${paramIndex++}`);
                    values.push(requestBody.summary);
                }
                if (requestBody.description !== undefined) {
                    setClauses.push(`description = $${paramIndex++}`);
                    values.push(requestBody.description);
                }
                if (requestBody.location !== undefined) {
                    setClauses.push(`location = $${paramIndex++}`);
                    values.push(requestBody.location);
                }
                if (requestBody.start?.dateTime !== undefined) {
                    setClauses.push(`start_datetime = $${paramIndex++}`);
                    values.push(resolveDateTime(requestBody.start.dateTime, requestBody.start.timeZone ?? null));
                }
                if (requestBody.start?.timeZone !== undefined) {
                    setClauses.push(`start_timezone = $${paramIndex++}`);
                    values.push(requestBody.start.timeZone);
                }
                if (requestBody.end?.dateTime !== undefined) {
                    setClauses.push(`end_datetime = $${paramIndex++}`);
                    values.push(resolveDateTime(requestBody.end.dateTime, requestBody.end.timeZone ?? null));
                }
                if (requestBody.end?.timeZone !== undefined) {
                    setClauses.push(`end_timezone = $${paramIndex++}`);
                    values.push(requestBody.end.timeZone);
                }
                if (requestBody.attendees !== undefined) {
                    setClauses.push(`attendees = $${paramIndex++}::jsonb`);
                    values.push(JSON.stringify(requestBody.attendees));
                }

                // Always update the updated timestamp
                setClauses.push(`updated = NOW()`);

                if (setClauses.length === 1) {
                    // Only the updated timestamp, no real changes; just fetch
                    const result = await pool.query(
                        `SELECT * FROM gcal.events WHERE id = $1`,
                        [eventId]
                    );
                    if (result.rows.length === 0) throw new Error(`Event not found: ${eventId}`);
                    return { data: formatEvent(result.rows[0]) };
                }

                values.push(eventId);
                const result = await pool.query(
                    `UPDATE gcal.events SET ${setClauses.join(', ')} WHERE id = $${paramIndex} RETURNING *`,
                    values
                );
                if (result.rows.length === 0) {
                    throw new Error(`Event not found: ${eventId}`);
                }
                return { data: formatEvent(result.rows[0]) };
            },

            async delete({ calendarId, eventId }) {
                const result = await pool.query(
                    `DELETE FROM gcal.events WHERE id = $1`,
                    [eventId]
                );
                if (result.rowCount === 0) {
                    throw new Error(`Event not found: ${eventId}`);
                }
                return { data: {} };
            },

            async list({ calendarId, timeMin, timeMax, maxResults, orderBy, singleEvents }) {
                const conditions: string[] = [];
                const values: any[] = [];
                let paramIndex = 1;

                if (timeMin) {
                    conditions.push(`start_datetime >= $${paramIndex++}`);
                    values.push(timeMin);
                }
                if (timeMax) {
                    conditions.push(`end_datetime <= $${paramIndex++}`);
                    values.push(timeMax);
                }

                const whereClause = conditions.length > 0
                    ? `WHERE ${conditions.join(' AND ')}`
                    : '';

                let orderClause = 'ORDER BY start_datetime ASC';
                if (orderBy === 'updated') {
                    orderClause = 'ORDER BY updated DESC';
                }

                const limitClause = maxResults ? `LIMIT $${paramIndex++}` : '';
                if (maxResults) {
                    values.push(maxResults);
                }

                const result = await pool.query(
                    `SELECT * FROM gcal.events ${whereClause} ${orderClause} ${limitClause}`,
                    values
                );
                return { data: { items: result.rows.map(formatEvent) } };
            },
        };
    }
}

export function createPool(): InstanceType<typeof Pool> {
    return new Pool({
        host: process.env.PG_HOST || 'localhost',
        port: parseInt(process.env.PG_PORT || '5432'),
        database: process.env.PG_DATABASE || 'cowork_gym',
        user: process.env.PG_USER || 'postgres',
        password: process.env.PG_PASSWORD || 'postgres',
    });
}
