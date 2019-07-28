import * as React from 'react';
import * as strings from 'ModernEventsWebPartStrings';
import FullCalendar from '@fullcalendar/react';
import dayGridPlugin from '@fullcalendar/daygrid';
import timeGridPlugin from '@fullcalendar/timegrid';
import listPlugin from '@fullcalendar/list';
import { IFullCalendarEvent } from "./IFullCalendarEvent";
import { ToolbarInput } from '@fullcalendar/core/types/input-types';
import { DisplayType } from './ENUMDisplayType';
import interactionPlugin from '@fullcalendar/interaction';
import {IInteraction,IDisplayOptions} from './CalendarApp';
export interface IEventCalendarProps {
  cbSelectEntry: any;
  cbUpdateEvents: any;
  cbDragDropEvent: any;
  cbNewEvent: any;
  events: IFullCalendarEvent[];
  displayType?: DisplayType;
  timeformat: string;
  interactions:IInteraction;
  displayOptions:IDisplayOptions;
}

export interface IEventCalendarState {
  events: IFullCalendarEvent[];
  displayType: DisplayType;
  currentDate: Date;
}

export class EventCalendar extends React.Component<IEventCalendarProps, IEventCalendarState> {
  private calRef: React.RefObject<FullCalendar>;

  constructor(props: IEventCalendarProps) {
    super(props);
    let displayType: DisplayType = (this.props.displayType) ? this.props.displayType : DisplayType.DayGrid;
    this.calRef = React.createRef();
    this.state = {
      events: this.props.events,
      displayType: displayType,
      currentDate: new Date()
    };
  }

  public getDisplayDate(): Date {
    return this.state.currentDate;
  }

  public getCurrentEventList(): IFullCalendarEvent[] {
    return this.state.events;
  }

  public componentWillReceiveProps(nextProps: IEventCalendarProps) {
    this.setState({
      displayType: (this.state.displayType != nextProps.displayType && nextProps.displayType) ? nextProps.displayType : this.state.displayType,
      events: (this.state.events != nextProps.events) ? nextProps.events : this.state.events,
    });
  }

  public render(): React.ReactElement<IEventCalendarProps> {

    let btn = {
      prevMonth: {
        text: '<',
        click: this._navigatePrev.bind(this)
      },
      nextMonth: {
        text: '>',
        click: this._navigateNext.bind(this)
      },
      todayCustom: {
        text: 'today',
        click: this._navigateToday.bind(this)
      }
    };
    let header: ToolbarInput = {
      left: "prevMonth",
      center: "title",
      right: "todayCustom nextMonth"

    };
    let defaultView = "dayGridMonth";
    let plugins = [interactionPlugin, timeGridPlugin];
    switch (this.state.displayType) {
      case DisplayType.DayGrid:
        defaultView = "dayGridMonth";
        plugins = [interactionPlugin, dayGridPlugin];
        break;
      case DisplayType.TimeGrid:
        defaultView = "timeGridWeek";
        plugins = [interactionPlugin, timeGridPlugin];
        break;
      case DisplayType.ListGrid:
        defaultView = "listWeek";
        plugins = [interactionPlugin, listPlugin];
        break;
      default:
        defaultView = "dayGridMonth";
        plugins = [interactionPlugin, dayGridPlugin];
        break;
    }


    return (
<<<<<<< HEAD
          <FullCalendar
            ref={this.calRef}
            defaultView={defaultView}
            plugins={plugins}
            themeSystem="standard"
            events={this.state.events}
            defaultDate={this.state.currentDate?this.state.currentDate:new Date()}
            eventClick={this._eventClick.bind(this)}
            eventMouseEnter={this._eventMouseEnter.bind(this)}
            eventMouseLeave={this._eventMouseLeave.bind(this)}
            customButtons={btn}
            header={header}
            dateClick={this._dateClick.bind(this)}
            slotLabelFormat={ {
              hour12: (this.props.timeformat=='12h')?true:false,
              hour: '2-digit',
              minute: '2-digit',
            }}
            eventTimeFormat={ {
              hour12: (this.props.timeformat=='12h')?true:false,
              hour: '2-digit',
              minute: '2-digit',
              second: '2-digit'
            }}
          />
=======
      <FullCalendar
        allDayText={strings.LabelAllDay}
        ref={this.calRef}
        defaultView={defaultView}
        plugins={plugins}
        themeSystem="standard"
        events={this.state.events}
        defaultDate={this.state.currentDate ? this.state.currentDate : new Date()}
        eventClick={this._eventClick.bind(this)}
        eventMouseEnter={this._eventMouseEnter.bind(this)}
        eventMouseLeave={this._eventMouseLeave.bind(this)}
        customButtons={btn}
        header={header}
        dateClick={this._dateClick.bind(this)}
        editable={this.props.interactions.dragAndDrop}
        eventDurationEditable={false}
        eventDragStart={this._eventDragtStart.bind(this)}
        eventDragStop={this._eventDragStop.bind(this)}
        eventDrop={this._eventDrop.bind(this)}
        firstDay={parseInt(this.props.displayOptions.weekStartsAt)}
        slotLabelFormat={{
          hour12: (this.props.timeformat == '12h') ? true : false,
          hour: '2-digit',
          minute: '2-digit',
        }}
        eventTimeFormat={{
          hour12: (this.props.timeformat == '12h') ? true : false,
          hour: '2-digit',
          minute: '2-digit',
          second: '2-digit'
        }}
      />
>>>>>>> Version-2
    );
  }
  private _eventDragtStart(event: any) {
  }
  private _eventDragStop(event: any) {

  }
  private _eventDrop(eventData: any) {
    if(!this.props.interactions.dragAndDrop){
      console.log("Drag&Drop Support disabled in config");
      return;
    }
    this.props.cbDragDropEvent(eventData.event);
  }
  private _navigateToday() {
    let calendarApi = this.calRef.current.getApi();
    calendarApi.today();
    this.props.cbUpdateEvents(this.state.displayType, calendarApi.getDate()).then((events) => {
      this.setState({
        events: events
      });
    });

  }
  private _dateClick(parms: any): void {
    this.props.cbNewEvent(parms.dateStr);
  }

  private _navigateNext(parms: any): void {
    let calendarApi = this.calRef.current.getApi();
    calendarApi.next();

    this.props.cbUpdateEvents(this.state.displayType, calendarApi.getDate()).then((events) => {
      this.setState({
        events: events
      });
    });

  }

  private _navigatePrev(parms: any): void {
    let calendarApi = this.calRef.current.getApi();
    calendarApi.prev();
    this.props.cbUpdateEvents(this.state.displayType, calendarApi.getDate()).then((events) => {
      this.setState({
        events: events
      });
    });

  }


  private _eventClick(parms: any) {
    let event: IFullCalendarEvent = parms.event;
    this.props.cbSelectEntry(event);
    return true;
  }

  private _eventMouseEnter(parms: any) {
    let event: IFullCalendarEvent = parms.event;
    return true;
  }
  private _eventMouseLeave(parms: any) {
    let event: IFullCalendarEvent = parms.event;
    return true;
  }
}
