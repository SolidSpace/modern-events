export interface IFullCalendarEvent{
  id:number;
  allDay:boolean;
  title:string;
  start:string;
  end:string;
  url?:string;
  classNames?:string[];
  editable?:boolean;
  extendedProps?:any;
}
