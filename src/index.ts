import * as dotenv from "dotenv";
import { OutlookClient } from "./clients/outlook.js";
import { GoogleCalendarClient } from "./clients/google.js";

dotenv.config();

async function getOutlookEvents() {
    const outlookClient = new OutlookClient(
        process.env.OUTLOOK_ICS_URL || ""
    );

    const events = await outlookClient.getEvents();
    console.log(`got ${events.length} events from Outlook`);
}

async function getGoogleEvents() {
    const googleClient = new GoogleCalendarClient(
        process.env.GOOGLE_CAL_ID || "",
        process.env.GOOGLE_CAL_KEY || ""
    );
    await googleClient.listEvents();
}

/*
getOutlookEvents().catch((error) => {
    console.error("Error during Outlook fetch:", error);
});
*/

getGoogleEvents().catch((error) => {
    console.error("Error during Google fetch:", error);
});
