import { ISPEvent } from "../components/ISPEvent";
import { IFullCalendarEvent } from "../components/IFullCalendarEvent";
import * as moment from 'moment';

export class EventConverter {

  public static getSPEvent(event: IFullCalendarEvent): ISPEvent {
    let spEvent = {
      Id:event.id,
      EventDate: event.start,
      EndDate: event.end,
      Title: event.title,
      fAllDayEvent: event.allDay,
      Description: event.extendedProps.description,
      Location: event.extendedProps.location,
      Category: event.extendedProps.category
    };

    if (event.allDay) {
      spEvent.EventDate = moment(event.start).set({ h: 0, m: 0 }).format("YYYY-MM-DDTHH:mm:ss");
      (event.end==null || !event.end)?spEvent.EndDate = moment(event.start).set({ h: 23, m: 59 }).format("YYYY-MM-DDTHH:mm:ss"):spEvent.EndDate = moment(event.end).set({ h: 23, m: 59 }).format("YYYY-MM-DDTHH:mm:ss");
    } else {
      spEvent.EventDate = moment(event.start).format("YYYY-MM-DDTHH:mm:ss");
      spEvent.EndDate = moment(event.end).format("YYYY-MM-DDTHH:mm:ss");
    }
     return spEvent;
  }

  public static getFCEvent(event: ISPEvent): IFullCalendarEvent {
    return {
      title: event.Title,
      id: event.Id,
      start: event.EventDate,
      end: event.EndDate,
      allDay: event.fAllDayEvent,
      extendedProps: {
        location: event.Location,
        description: event.Description,
        category: event.Category
      }
    };
  }

}
