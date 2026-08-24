import type { IcsEvent } from "./types.d.ts";

export class OutlookClient {
    constructor(
        private icsUrl: string, 
    ) {}
    
    private unfoldLines(icsText: string): string[] {
        return icsText
            .replace(/\r\n[ \t]/g, '')
            .replace(/\n[ \t]/g, '')
            .split(/\r?\n/);
    }

    private parseIcsDate(val: string) {
        const match = val.match(/^(\d{4})(\d{2})(\d{2})(?:T(\d{2})(\d{2})(\d{2})(Z)?)?$/);
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

    async getEvents(): Promise<IcsEvent[]> {
        const response = await fetch(this.icsUrl);

        if (!response.ok) {
            throw new Error(`Failed to download ICS feed: HTTP ${response.status}`);
        }

        const icsText = await response.text();
        const lines = this.unfoldLines(icsText);

        const events: IcsEvent[] = [];
        let currentEvent: Partial<IcsEvent> | null = null;

        for (const line of lines) {
            if (line.startsWith('BEGIN:VEVENT')) {
                currentEvent = {};
            } else if (line.startsWith('END:VEVENT')) {
                if (currentEvent && currentEvent.start) {
                    events.push(currentEvent as IcsEvent);
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
                            currentEvent.summary = val.trim();
                            break;
                        case 'DESCRIPTION':
                            currentEvent.description = this.unescapeIcsText(val.trim());
                            break;
                        case 'UID':
                            currentEvent.uid = val.trim();
                            break;
                    }
                }
            }
        }

        return events.filter(ev => ev.summary && !ev.summary.toLowerCase().startsWith('canceled:'));
    }
}