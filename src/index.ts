import { syncCalendar } from "./service.js";

console.log("Calendar sync service stated");
syncCalendar().catch(err => {
    console.error("Error syncing calendar:", err);
});
