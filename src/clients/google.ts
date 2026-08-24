import { google } from "googleapis";
import path from "path";

const SCOPES = ['https://www.googleapis.com/auth/calendar.events'];

export class GoogleCalendarClient {
    private credentialsPath: string;

    constructor(
        private readonly calendarId: string,
        keyFileName: string,
    ) {
        this.credentialsPath = path.join(process.cwd(), 'keys/' + keyFileName);
    }

    public async listEvents() {
        const auth = new google.auth.GoogleAuth({
            keyFile: this.credentialsPath,
            scopes: SCOPES,
        });

        const calendar = google.calendar({version: 'v3', auth});
        const result = await calendar.events.list({
            calendarId: this.calendarId,
            timeMin: new Date().toISOString(),
            maxResults: 10,
            singleEvents: true,
            orderBy: 'startTime',
        });

        const events = result.data.items;
        if (!events || events.length === 0) {
            console.log('No upcoming events found.');
            return;
        }
        console.log('Upcoming 10 events:');

        for (const event of events) {
            const start = event.start?.dateTime ?? event.start?.date;
            console.log(`${start} - ${event.summary}`);
        }
    }
}