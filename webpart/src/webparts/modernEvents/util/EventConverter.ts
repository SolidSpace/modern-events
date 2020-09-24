import { ISPEvent } from "../components/ISPEvent";
import { IFullCalendarEvent } from "../components/IFullCalendarEvent";
import * as moment from 'moment';
import { IFieldMap } from "../components/IFieldMap";
export interface IEventDelta {
  days: number;
  milliseconds: number;
  months: number;
  years: number;
}

export class EventConverter {

  public static getSPEvent(event: IFullCalendarEvent): ISPEvent {
    let spEvent =   {
      Id: event.id,
      EventDate: event.start,
      EndDate: EventConverter.tidyEndDate(event.start, event.end, false, 1),
      Title: event.title,
      fAllDayEvent: event.allDay,
      Description: event.extendedProps.description,
      Location: event.extendedProps.location,
      Category: event.extendedProps.category
    };

    if (event.allDay) {
      spEvent.EventDate = moment(event.start).set({ h: 0, m: 0 }).format("YYYY-MM-DDTHH:mm:ss");
      spEvent.EndDate = (event.end == null) ? moment(event.start).set({ h: 23, m: 59 }).format("YYYY-MM-DDTHH:mm:ss") : spEvent.EndDate;
    } else {
      spEvent.EventDate = moment(event.start).format("YYYY-MM-DDTHH:mm:ss");
      spEvent.EndDate = moment(event.end).format("YYYY-MM-DDTHH:mm:ss");
    }
    return spEvent;
  }

  public static getCustomEvent(event: ISPEvent, fieldMap: IFieldMap): any {
    if (fieldMap.isDefaultSchema) {
      return event;
    }else{
      let spEvent = {};
      spEvent['Id']=event.Id;
      spEvent[fieldMap.Title]=event.Title;
      spEvent[fieldMap.EventDate]=event.EventDate;
      spEvent[fieldMap.EndDate]=event.EndDate;
      spEvent[fieldMap.fAllDayEvent]=event.fAllDayEvent;
      spEvent[fieldMap.Location]=event.Location;
      spEvent[fieldMap.Category]=event.Category;
      spEvent[fieldMap.Description]=event.Description;
      return spEvent;
    }

  }

  public static getFCEvent(event: ISPEvent, fieldMap: IFieldMap): IFullCalendarEvent {
    let fcEvent: IFullCalendarEvent;
    if (fieldMap.isDefaultSchema) {
      fcEvent = {
        title: event.Title,
        id: event.Id,
        start: event.EventDate,
        end: EventConverter.tidyEndDate(event.EventDate, event.EndDate, true, 1),
        allDay: event.fAllDayEvent,
        extendedProps: {
          location: event.Location,
          description: event.Description,
          category: event.Category
        }
      };
    } else {
      fcEvent = {
        title: event[fieldMap.Title],
        id: event.Id,
        start: event[fieldMap.EventDate],
        end: EventConverter.tidyEndDate(event[fieldMap.EventDate], event[fieldMap.EndDate], true, 1),
        allDay: (event[fieldMap.fAllDayEvent] == "1") ? true : false,
        extendedProps: {
          location: event[fieldMap.Location],
          description: event[fieldMap.Description],
          category: event[fieldMap.Category]
        }
      };

    }
    return fcEvent;
  }
  /**
   * Fix End Date for multiple All Day Events. SharePoint displays last Day correctly but in
   * Fullcalendar the Event Ends before the last day. See documentation here:
   * Issue: https://github.com/SolidSpace/modern-events/issues/1
   *
   * https://fullcalendar.io/docs/event-object
   * end:
   * This value is exclusive! I repeat, this value is exclusive!!!
   * An event with the end of 2018-09-03 will appear to span through the 2nd of the month,
   * but will end before the start of the 3rd of the month.
   */
  public static tidyEndDate(startDate: string, endDate: string, addDays: boolean, days: number): string {
    if (moment(endDate, "YYYY-MM-DDTHH:mm:ss").set({ h: 0, m: 0 }).isAfter(moment(startDate, "YYYY-MM-DDTHH:mm:ss").set({ h: 0, m: 0 }))) {
      return (addDays) ? moment(endDate, "YYYY-MM-DDTHH:mm:ss").add(days, 'days').format("YYYY-MM-DDTHH:mm:ss") : moment(endDate, "YYYY-MM-DDTHH:mm:ss").subtract(days, 'days').format("YYYY-MM-DDTHH:mm:ss");
    } else {
      return endDate;
    }
  }

}
