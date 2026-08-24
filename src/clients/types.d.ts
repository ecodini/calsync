export interface IcsEvent {
  uid?: string;
  summary?: string;
  description?: string;
  start: Date;
  end?: Date;
}
