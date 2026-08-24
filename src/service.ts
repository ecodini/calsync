import { GoogleCalendarClient } from "./clients/google.js";
import { OutlookClient } from "./clients/outlook.js";
import * as dotenv from "dotenv";
import * as fs from "fs";

dotenv.config();

const outlookClient = new OutlookClient(
    process.env.OUTLOOK_ICS_URL || ""
);

const googleClient = new GoogleCalendarClient(
    process.env.GOOGLE_CAL_ID || "",
    process.env.GOOGLE_CAL_KEY || ""
);

const rawData = JSON.parse(fs.readFileSync('./windowsZones.json', 'utf-8'));
const mapTimezones = rawData.supplemental.windowsZones.mapTimezones;

const windowsToIana: Record<string, string> = {};

for (const entry of mapTimezones) {
    const zone = entry.mapZone;
    
    if (zone._territory === '001') {
        const primaryIana = zone._type.split(' ')[0];
        
        windowsToIana[zone._other] = primaryIana;
    }
}

const DAYS_TO_SYNC = 90;

const windowStart = new Date(Date.now() - (DAYS_TO_SYNC * 24 * 60 * 60 * 1000));
const windowEnd = new Date(Date.now() + (DAYS_TO_SYNC * 24 * 60 * 60 * 1000));

export async function syncCalendar() {
    console.log("Running sync at ", new Date().toISOString());
    const outlookCalendarItems = await outlookClient.getEvents();
    console.log(`Fetched ${outlookCalendarItems.length} events from Outlook calendar.`);3

    const googleIDs = new Set<string>();

    for (const outlookEvent of outlookCalendarItems) {
        if (!outlookEvent.uid) {
            continue;
        }

        if (outlookEvent.start < windowStart) {
            continue
        }

        if (outlookEvent.start > windowEnd) {
            continue;
        }

        console.log(`Processing event: ${outlookEvent.summary} (UID: ${outlookEvent.uid})`);

        const newId = GoogleCalendarClient.generateGoogleId(outlookEvent.uid);

        googleIDs.add(newId);

        const fallbackEndString = outlookEvent.endString || outlookEvent.startString;

        await googleClient.upsertEvent(newId, {
            summary: outlookEvent.summary ?? 'Busy',
            description: outlookEvent.description ?? '',
            start: {
                dateTime: outlookEvent.startString,
                timeZone: windowsToIana[outlookEvent.timeZone] || outlookEvent.timeZone,
            },
            end: {
                dateTime: fallbackEndString,
                timeZone: windowsToIana[outlookEvent.timeZone] || outlookEvent.timeZone,
            },
            extendedProperties: {
            private: { syncedFromIcs: 'true' }
        }
        });
    }

    await googleClient.reconcileEvents(googleIDs, windowStart, windowEnd);
}