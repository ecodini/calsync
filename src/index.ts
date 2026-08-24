import * as dotenv from "dotenv";
import { OutlookClient } from "./clients/outlook.js";

dotenv.config();

async function getOutlookEvents() {
    const outlookClient = new OutlookClient(
        process.env.OUTLOOK_ICS_URL || ""
    );

    const events = await outlookClient.getEvents();
    console.log(`got ${events.length} events from Outlook`);
}

getOutlookEvents().catch((error) => {
    console.error("Error during Outlook fetch:", error);
});
