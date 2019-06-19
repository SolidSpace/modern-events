import * as React from 'react';
import FullCalendar from '@fullcalendar/react';
import dayGridPlugin from '@fullcalendar/daygrid';
import timeGridPlugin from '@fullcalendar/timegrid';
import listPlugin from '@fullcalendar/list';
import { IFullCalendarEvent } from "./IFullCalendarEvent";
import {  ToolbarInput } from '@fullcalendar/core/types/input-types';
import {DisplayType} from './ENUMDisplayType';
import interactionPlugin from '@fullcalendar/interaction';

export interface IEventCalendarProps {
  cbSelectEntry:any;
  cbUpdateEvents:any;
  events: IFullCalendarEvent[];
  displayType?:DisplayType;
  timeformat:string;
}

export interface IEventCalendarState {
  events: IFullCalendarEvent[];
  displayType:DisplayType;
  currentDate:Date;
}

export class EventCalendar extends React.Component<IEventCalendarProps, IEventCalendarState> {
  private calRef: React.RefObject<FullCalendar>;

  constructor(props: IEventCalendarProps) {
    super(props);
    let displayType:DisplayType =(this.props.displayType)? this.props.displayType : DisplayType.DayGrid;
    this.calRef = React.createRef();
    this.state = {
      events:this.props.events,
      displayType:displayType,
      currentDate:new Date()
    };
  }

public getDisplayDate():Date{
  return this.state.currentDate;
}

public getCurrentEventList():IFullCalendarEvent[]{
  return this.state.events;
}

public componentWillReceiveProps(nextProps:IEventCalendarProps){
  this.setState({
    displayType:(this.state.displayType!=nextProps.displayType && nextProps.displayType)?nextProps.displayType:this.state.displayType,
    events:(this.state.events!=nextProps.events)?nextProps.events:this.state.events,
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
      todayCustom:{
        text:'today',
        click: this._navigateToday.bind(this)
      }
    };
    let header: ToolbarInput = {
      left: "prevMonth",
      center: "title",
      right: "todayCustom nextMonth"

    };
    let defaultView="dayGridMonth";
    let plugins=[timeGridPlugin];
    switch (this.state.displayType) {
      case DisplayType.DayGrid:
        defaultView="dayGridMonth";
        plugins=[dayGridPlugin];
      break;
      case DisplayType.TimeGrid:
        defaultView="timeGridWeek";
        plugins=[timeGridPlugin];
      break;
      case DisplayType.ListGrid:
        defaultView="listWeek";
        plugins=[listPlugin];
      break;
      default:
        defaultView="dayGridMonth";
        plugins=[dayGridPlugin];
      break;
    }


    return (
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
            eventTimeFormat={ {
              hour12: (this.props.timeformat=='12h')?true:false,
              hour: '2-digit',
              minute: '2-digit',
              second: '2-digit'
            }}
          />
    );

  }
  private _navigateToday(){
    let calendarApi = this.calRef.current.getApi();
    calendarApi.today();
    this.props.cbUpdateEvents(this.state.displayType,calendarApi.getDate()).then((events)=>{
      this.setState({
        events:events
      });
    });

  }
  private _dateClick(parms: any): void {
    console.log("n  ext");
  }

  private _navigateNext(parms: any): void {
    let calendarApi = this.calRef.current.getApi();
    calendarApi.next();

    this.props.cbUpdateEvents(this.state.displayType,calendarApi.getDate()).then((events)=>{
      this.setState({
        events:events
      });
    });

  }

  private _navigatePrev(parms: any): void {
    let calendarApi = this.calRef.current.getApi();
    calendarApi.prev();
    this.props.cbUpdateEvents(this.state.displayType,calendarApi.getDate()).then((events)=>{
      this.setState({
        events:events
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
