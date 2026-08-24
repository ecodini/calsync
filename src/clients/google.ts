import { google, calendar_v3 } from "googleapis";
import { GoogleAuth } from "google-auth-library";
import crypto from "node:crypto";
import path from "path";

const SCOPES = ['https://www.googleapis.com/auth/calendar.events'];

export class GoogleCalendarClient {
    private readonly credentialsPath: string;
    private auth: GoogleAuth;

    constructor(
        private readonly calendarId: string,
        keyFileName: string,
    ) {
        this.credentialsPath = path.join(process.cwd(), 'keys/' + keyFileName);

        this.auth = new GoogleAuth({
            keyFile: this.credentialsPath,
            scopes: SCOPES,
        });
    }

    private get client(): calendar_v3.Calendar {
        return google.calendar({ version: 'v3', auth: this.auth });
    }

    public async listEvents() {
        const result = await this.client.events.list({
            calendarId: this.calendarId,
            timeMin: new Date().toISOString(),
            maxResults: 10,
            singleEvents: true,
            orderBy: 'startTime',
        });

        const events = result.data.items;
        if (!events || events.length === 0) {
            console.log('No upcoming events found.');
            return [];
        }
        
        console.log(`Found ${events.length} upcoming events.`);
        return events;
    }

    public async getEvent(eventId: string) {
        try {
            const result = await this.client.events.get({
                calendarId: this.calendarId,
                eventId: eventId,
            });
            return result.data;
        } catch (error: any) {
            if (error.code === 404) {
                return null; // Return null if not found instead of crashing
            }
            throw error;
        }
    }

    public async createEvent(eventBody: calendar_v3.Schema$Event) {
        const result = await this.client.events.insert({
            calendarId: this.calendarId,
            requestBody: eventBody,
        });
        return result.data;
    }

    public async updateEvent(eventId: string, eventBody: calendar_v3.Schema$Event) {
        const result = await this.client.events.update({
            calendarId: this.calendarId,
            eventId: eventId,
            requestBody: eventBody,
        });
        return result.data;
    }

    public async deleteEvent(eventId: string) {
        try {
            await this.client.events.delete({
                calendarId: this.calendarId,
                eventId: eventId,
            });
            console.log(`Deleted event: ${eventId}`);
        } catch (error: any) {
            // Ignore 404s or 410s if the event is already gone
            if (error.code === 404 || error.code === 410) {
                console.log(`Event ${eventId} already deleted or not found.`);
                return;
            }
            throw error;
        }
    }

    public async upsertEvent(eventId: string, eventBody: calendar_v3.Schema$Event) {
        eventBody.id = eventId;

        try {
            const result = await this.client.events.insert({
                calendarId: this.calendarId,
                requestBody: eventBody,
            });
            console.log(`Inserted new event: ${eventId}`);
            return result.data;
            
        } catch (error: any) {
            if (error.code === 409) {
                console.log(`Event ${eventId} exists. Updating...`);
                
                const result = await this.client.events.update({
                    calendarId: this.calendarId,
                    eventId: eventId,
                    requestBody: eventBody,
                });
                return result.data;
            }
            
            throw error; 
        }
    }

    public async reconcileEvents(
        expectedGoogleIds: Set<string>, 
        timeMin: Date, 
        timeMax: Date
    ) {
        console.log('Starting reconciliation process...');

        try {
            const result = await this.client.events.list({
                calendarId: this.calendarId,
                timeMin: timeMin.toISOString(),
                timeMax: timeMax.toISOString(),
                privateExtendedProperty: ['syncedFromIcs=true'],
                singleEvents: true,
                maxResults: 2500,
            });

            const existingEvents = result.data.items || [];

            let deletedCount = 0;

            for (const gEvent of existingEvents) {
                if (gEvent.id && !expectedGoogleIds.has(gEvent.id)) {
                    console.log(`Orphaned event found: "${gEvent.summary}". Deleting...`);
                    await this.deleteEvent(gEvent.id);
                    deletedCount++;
                }
            }

            console.log(`Reconciliation complete. Deleted ${deletedCount} cancelled/moved events.`);

        } catch (error) {
            console.error('Error during reconciliation:', error);
            throw error;
        }
    }

    public static generateGoogleId(outlookUid: string): string {
        return crypto
            .createHash('sha256')
            .update(outlookUid)
            .digest('hex');
    }
}