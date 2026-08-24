export interface IcsEvent {
  uid?: string;
  summary?: string;
  description?: string;
  start: Date;
  end?: Date;
  startString: string; 
  endString?: string;
  timeZone: string;
  
  rrule?: string;
  exdates?: string[];
  recurrenceId?: string;
}

export class OutlookClient {
    constructor(private icsUrl: string) {}
    
    private unfoldLines(icsText: string): string[] {
        return icsText
            .replace(/\r\n[ \t]/g, '')
            .replace(/\n[ \t]/g, '')
            .split(/\r?\n/);
    }

    private parseIcsDate(val: string) {
        const dateValue = val.includes(':') ? val.split(':').pop()! : val;
        
        const match = dateValue.match(/^(\d{4})(\d{2})(\d{2})(?:T(\d{2})(\d{2})(\d{2})(Z)?)?$/);
        if (!match) return null;

        const [, year, month, day, hours = '00', minutes = '00', seconds = '00', isUtc] = match;
        
        const formattedString = `${year}-${month}-${day}T${hours}:${minutes}:${seconds}`;
        
        let date: Date;
        if (isUtc) {
            date = new Date(Date.UTC(+year!, +month! - 1, +day!, +hours, +minutes, +seconds));
        } else {
            date = new Date(+year!, +month! - 1, +day!, +hours, +minutes, +seconds);
        }

        return { date, formattedString, isUtc: !!isUtc };
    }

    private unescapeIcsText(text: string): string {
        return text
            .replace(/\\n/gi, '\n')
            .replace(/\\;/g, ';')
            .replace(/\\,/g, ',')
            .replace(/\\\\/g, '\\');
    }

    private shiftDateString(originalString: string, newDate: Date): string {
        const yyyy = newDate.getFullYear().toString();
        const mm = String(newDate.getMonth() + 1).padStart(2, '0');
        const dd = String(newDate.getDate()).padStart(2, '0');
        const timePart = originalString.split('T')[1] || '00:00:00';
        return `${yyyy}-${mm}-${dd}T${timePart}`;
    }

    private parseRrule(rrule: string) {
        const rule: any = { interval: 1 };
        const parts = rrule.split(';');
        for (const part of parts) {
            const [key, val] = part.split('=');
            if (key === 'FREQ') rule.freq = val;
            if (key === 'INTERVAL') rule.interval = parseInt(val!, 10);
            if (key === 'UNTIL') {
                const parsed = this.parseIcsDate(val!);
                if (parsed) rule.until = parsed.date;
            }
            if (key === 'BYDAY') rule.byday = val!.split(',');
        }
        return rule;
    }

    async getEvents(): Promise<IcsEvent[]> {
        const response = await fetch(this.icsUrl);

        if (!response.ok) {
            throw new Error(`Failed to download ICS feed: HTTP ${response.status}`);
        }

        const icsText = await response.text();
        const lines = this.unfoldLines(icsText);

        const rawEvents: IcsEvent[] = [];
        let currentEvent: Partial<IcsEvent> | null = null;

        // --- 1. PARSING LOOP ---
        for (const line of lines) {
            if (line.startsWith('BEGIN:VEVENT')) {
                currentEvent = {};
            } else if (line.startsWith('END:VEVENT')) {
                if (currentEvent && currentEvent.start) {
                    rawEvents.push(currentEvent as IcsEvent);
                }
                currentEvent = null;
            } else if (currentEvent) {
                const colonIdx = line.indexOf(':');

                if (colonIdx !== -1) {
                    const keyPart = line.substring(0, colonIdx);
                    const val = line.substring(colonIdx + 1);
                    const propName = keyPart.split(';')[0]?.trim() ?? '';

                    const tzidMatch = keyPart.match(/TZID=([^;]+)/);
                    const tzid = tzidMatch && tzidMatch.length > 1 ? tzidMatch[1]!.replace(/['"]/g, '') : null;

                    switch (propName) {
                        case 'DTSTART': {
                            const parsed = this.parseIcsDate(val.trim());
                            if (parsed) {
                                currentEvent.start = parsed.date;
                                currentEvent.startString = parsed.formattedString;
                                currentEvent.timeZone = tzid || (parsed.isUtc ? 'UTC' : 'UTC'); 
                            }
                            break;
                        }
                        case 'DTEND': {
                            const parsed = this.parseIcsDate(val.trim());
                            if (parsed) {
                                currentEvent.end = parsed.date;
                                currentEvent.endString = parsed.formattedString;
                            }
                            break;
                        }
                        case 'SUMMARY':
                            currentEvent.summary = this.unescapeIcsText(val.trim());
                            break;
                        case 'DESCRIPTION':
                            currentEvent.description = this.unescapeIcsText(val.trim());
                            break;
                        case 'UID':
                            currentEvent.uid = val.trim();
                            break;
                        case 'RRULE':
                            currentEvent.rrule = val.trim();
                            break;
                        case 'EXDATE': {
                            const parsed = this.parseIcsDate(val.trim());
                            if (parsed) {
                                currentEvent.exdates = currentEvent.exdates || [];
                                currentEvent.exdates.push(parsed.formattedString);
                            }
                            break;
                        }
                        case 'RECURRENCE-ID': {
                            const parsed = this.parseIcsDate(val.trim());
                            if (parsed) {
                                currentEvent.recurrenceId = parsed.formattedString;
                            }
                            break;
                        }
                    }
                }
            }
        }

        // --- 2. SORT INTO BUCKETS ---
        const activeEvents = rawEvents.filter(ev => ev.summary && !ev.summary.toLowerCase().startsWith('canceled:'));
        
        const normalEvents: IcsEvent[] = [];
        const masterEvents: IcsEvent[] = [];
        const overrides = new Map<string, IcsEvent>();

        for (const ev of activeEvents) {
            if (ev.recurrenceId) {
                overrides.set(`${ev.uid}_${ev.recurrenceId}`, ev);
            } else if (ev.rrule) {
                masterEvents.push(ev);
            } else {
                normalEvents.push(ev);
            }
        }

        // --- 3. FLATTEN RECURRENCES ---
        const flattenedEvents: IcsEvent[] = [];
        const dayStrings = ['SU', 'MO', 'TU', 'WE', 'TH', 'FR', 'SA'];
        
        // Define the 3-month (90-day) limit from today
        const today = new Date();
        const maxRecurrenceLimit = new Date(today.getTime() + 90 * 24 * 60 * 60 * 1000);

        for (const master of masterEvents) {
            const rule = this.parseRrule(master.rrule!);
            
            // Limit is rule.until OR 3 months from today, whichever comes first
            const limitDate = (rule.until && rule.until < maxRecurrenceLimit) ? rule.until : maxRecurrenceLimit;
            
            let currentDate = new Date(master.start);
            
            let loopGuard = 0;
            while (currentDate <= limitDate && loopGuard < 2500) {
                let isValid = false;
                const dayStr = dayStrings[currentDate.getDay()];

                if (rule.freq === 'DAILY') {
                    const daysElapsed = Math.round((currentDate.getTime() - master.start.getTime()) / 86400000);
                    if (daysElapsed % rule.interval === 0) isValid = true;
                } else if (rule.freq === 'WEEKLY') {
                    if (rule.byday && rule.byday.includes(dayStr)) {
                        const daysElapsed = Math.round((currentDate.getTime() - master.start.getTime()) / 86400000);
                        const weeksElapsed = Math.floor(daysElapsed / 7);
                        if (weeksElapsed % rule.interval === 0) isValid = true;
                    }
                } else if (rule.freq === 'MONTHLY' || rule.freq === 'YEARLY') {
                    if (currentDate.getDate() === master.start.getDate()) isValid = true;
                }

                if (isValid) {
                    const instanceStartStr = this.shiftDateString(master.startString, currentDate);
                    
                    if (!master.exdates?.includes(instanceStartStr)) {
                        const overrideKey = `${master.uid}_${instanceStartStr}`;
                        
                        if (overrides.has(overrideKey)) {
                            flattenedEvents.push(overrides.get(overrideKey)!);
                        } else {
                            const durationMs = (master.end?.getTime() ?? master.start.getTime()) - master.start.getTime();
                            const instanceEnd = new Date(currentDate.getTime() + durationMs);
                            
                            flattenedEvents.push({
                                ...master,
                                uid: `${master.uid}_${instanceStartStr}`, 
                                start: new Date(currentDate),
                                end: instanceEnd,
                                startString: instanceStartStr,
                                endString: master.endString ? this.shiftDateString(master.endString, instanceEnd) : undefined,
                            });
                        }
                    }
                }
                
                currentDate.setDate(currentDate.getDate() + 1);
                loopGuard++;
            }
        }

        // --- 4. RETURN EVERYTHING ---
        return [...normalEvents, ...flattenedEvents];
    }
}