import { syncCalendar } from "./service.js";
import cron from 'node-cron';

console.log("Calendar sync service stated");

syncCalendar().catch(err => {
    console.error("Error syncing calendar:", err);
});

// run every 2 hours from 8am to 8pm, Monday to Friday
cron.schedule('0 8-20/2 * * 1-5', () => {
    syncCalendar().catch(err => {
        console.error("Error syncing calendar:", err);
    });
}, { noOverlap: true, timezone: process.env.TZ || 'America/Argentina/Buenos_Aires' });
