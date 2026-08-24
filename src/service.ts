import { GoogleCalendarClient } from "./clients/google.js";
import { OutlookClient } from "./clients/outlook.js";
import * as dotenv from "dotenv";


dotenv.config();

const outlookClient = new OutlookClient(
    process.env.OUTLOOK_ICS_URL || ""
);

const googleClient = new GoogleCalendarClient(
    process.env.GOOGLE_CAL_ID || "",
    process.env.GOOGLE_CAL_KEY || ""
);

export async function syncCalendar() {
    console.log("Running sync at ", new Date().toISOString());
    const outlookCalendarItems = await outlookClient.getEvents();
    console.log(`Fetched ${outlookCalendarItems.length} events from Outlook calendar.`);

    for (const outlookEvent of outlookCalendarItems) {
        if (!outlookEvent.uid) {
            continue;
        }

        if (outlookEvent.start < new Date()) {
            continue
        }

        if (outlookEvent.start > new Date(Date.now() + 1 * 24 * 60 * 60 * 1000)) {
            continue;
        }

        console.log(`Processing event: ${outlookEvent.summary} (UID: ${outlookEvent.uid})`);

        const newId = GoogleCalendarClient.generateGoogleId(outlookEvent.uid);

        await googleClient.upsertEvent(newId, {
            summary: outlookEvent.summary ?? 'Busy',
            description: outlookEvent.description ?? '',
            start: {
                dateTime: outlookEvent.start.toISOString(),
                timeZone: 'UTC',
            },
            end: {
                dateTime: (outlookEvent.end || new Date(outlookEvent.start.getTime() + 60 * 60 * 1000)).toISOString(),
                timeZone: 'UTC',
            },
            extendedProperties: {
                private: { syncedFromIcs: 'true' }
            }
        });
    }
}