import type { IcsEvent } from "./types.d.ts";

export class OutlookClient {
    constructor(
        private icsUrl: string, 
    ) {}
    
    _unfoldLines(icsText: string): string[] {
        return icsText
            .replace(/\r\n[ \t]/g, '')
            .replace(/\n[ \t]/g, '')
            .split(/\r?\n/);
    }

    _parseIcsDate(rawLine: string): Date | null {
        const dateValue = rawLine.includes(':') ? rawLine.split(':').pop() : rawLine;
        if (!dateValue) return null;

        const match = dateValue.match(/^(\d{4})(\d{2})(\d{2})(?:T(\d{2})(\d{2})(\d{2})(Z)?)?$/);
        if (!match) return null;

        const [, year, month, day, hours = '00', minutes = '00', seconds = '00', isUtc] = match;

        if (isUtc) {
            return new Date(Date.UTC(+year!, +month! - 1, +day!, +hours, +minutes, +seconds));
        }
        return new Date(+year!, +month! - 1, +day!, +hours, +minutes, +seconds);
    }

    async getEvents(): Promise<IcsEvent[]> {
        const response = await fetch(this.icsUrl);

        if (!response.ok) {
            throw new Error(`Failed to download ICS feed: HTTP ${response.status}`);
        }

        const icsText = await response.text();
        const lines = this._unfoldLines(icsText);

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

                    switch (propName) {
                        case 'DTSTART':
                            currentEvent.start = this._parseIcsDate(line)!;
                            break;
                        case 'DTEND':
                            currentEvent.end = this._parseIcsDate(line)!;
                            break;
                        case 'SUMMARY':
                            currentEvent.summary = val.trim();
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