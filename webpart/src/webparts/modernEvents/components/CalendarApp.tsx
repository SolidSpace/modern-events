import * as React from 'react';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { MockupData } from "./MockupData";
import { CalendarCommandbar } from './CalendarCommandbar';
import { DisplayType } from './ENUMDisplayType';
import { EventCalendar } from "./EventCalendar";
import { IFullCalendarEvent } from "./IFullCalendarEvent";
import { EventPanel } from './EventPanel';
import * as CamlBuilder from 'camljs';
import FullCalendar from '@fullcalendar/react';
import * as moment from 'moment';
import { PnPListConnector } from '../services/PnPListConnector';
import { SiteConnector } from '../services/SiteConnector';
import { EventConverter } from "../util/EventConverter";
import { ISPEvent } from './ISPEvent';
import { ItemUpdateResult, ItemAddResult } from '@pnp/sp';
import { ICBButtonVisibility } from "./ICBButtonVisibility";
import { IFieldMap } from   './IFieldMap';
import ModernEventsWebPart from '../ModernEventsWebPart';
export interface IInteraction{
  dateClickNew:boolean;
  dragAndDrop:boolean;
}

export interface IDisplayOptions{
  weekStartsAt:string;
}

export interface ICalendarAppProps {
  fieldMapping:IFieldMap;
  context: any;
  listName: string;
  remoteSiteUrl: string;
  relativeLibOrListUrl: string;
  displayType: DisplayType;
  commandBarButtonVisibility: ICBButtonVisibility;
  commandBarVisible:boolean;
  timeformat:string;
  interactions:IInteraction;
  displayOptions:IDisplayOptions;
  overlayListCollection?: any[];
  isListOverviewEnabled?: boolean;
}

export interface ICalendarAppState {
  commandbar?: JSX.Element;
  content?: JSX.Element | FullCalendar;
  panel?: JSX.Element;
  showPanel: boolean;
  displayType: DisplayType;
  categories?: any[];
}

export interface IListOverlayCfg{
  isEnabled:boolean;
  collection:any[];
}

export class CalendarApp extends React.Component<ICalendarAppProps, ICalendarAppState> {
  private eventCalRef: React.RefObject<EventCalendar>;

  constructor(props: ICalendarAppProps) {
    super(props);
    this.eventCalRef = React.createRef();
    this.state = {
      commandbar: <CalendarCommandbar
        cbListGrid={this._changeDisplayType.bind(this)}
        cbTimeGrid={this._changeDisplayType.bind(this)}
        cbWeekGrid={this._changeDisplayType.bind(this)}
        cbNewEntry={this._newEntry.bind(this)}
        buttonVisibiliy={this.props.commandBarButtonVisibility}
      ></CalendarCommandbar>,
      displayType: props.displayType,
      showPanel: false,
    };

  }

  public componentDidMount() {
    let displayDate: Date;
    try {
      displayDate = this.eventCalRef.current.getDisplayDate();
    } catch (e) {
    }
    displayDate = displayDate ? displayDate : new Date();

    if (this.props.remoteSiteUrl && this.props.relativeLibOrListUrl && this.props.listName) {
      this._queryEvents(this.props.displayType, displayDate).then((calEvents) => {

        this.setState({
          content: <EventCalendar
                      displayType={this.state.displayType}
                      ref={this.eventCalRef}
                      cbSelectEntry={this._selectedEntry.bind(this)}
                      cbUpdateEvents={this._queryEvents.bind(this)}
                      events={calEvents}
                      timeformat={this.props.timeformat}
                      cbNewEvent={this._newEntry.bind(this)}
                      cbDragDropEvent={this._dragDropUpdateEvent.bind(this)}
                      interactions={this.props.interactions}
                      displayOptions={this.props.displayOptions}
                      >
                      </EventCalendar>
        });
      });
    } else {
      this.setState({
        content: <EventCalendar
                    displayType={this.state.displayType}
                    ref={this.eventCalRef}
                    cbSelectEntry={this._selectedEntry.bind(this)}
                    cbUpdateEvents={this._queryEvents.bind(this)}
                    events={[]}
                    timeformat={this.props.timeformat}
                    cbNewEvent={this._newEntry.bind(this)}
                    cbDragDropEvent={this._dragDropUpdateEvent.bind(this)}
                    interactions={this.props.interactions}
                    displayOptions={this.props.displayOptions}
                    ></EventCalendar>
      });

    }

  }

  public componentWillReceiveProps(nextProps: ICalendarAppProps) {
    console.log(nextProps);
    let displayDate: Date = this.eventCalRef.current.getDisplayDate();
    displayDate = displayDate ? displayDate : new Date();

    if (
      (nextProps.remoteSiteUrl && nextProps.relativeLibOrListUrl && nextProps.listName) ||
      (nextProps.commandBarButtonVisibility != this.props.commandBarButtonVisibility)) {

      this._queryEvents(this.props.displayType, displayDate, nextProps).then((calEvents) => {

        this.setState({
          commandbar: <CalendarCommandbar
            cbListGrid={this._changeDisplayType.bind(this)}
            cbTimeGrid={this._changeDisplayType.bind(this)}
            cbWeekGrid={this._changeDisplayType.bind(this)}
            cbNewEntry={this._newEntry.bind(this)}
            buttonVisibiliy={nextProps.commandBarButtonVisibility}
          ></CalendarCommandbar>,
          content: <EventCalendar
                      ref={this.eventCalRef}
                      cbSelectEntry={this._selectedEntry.bind(this)}
                      cbUpdateEvents={this._queryEvents.bind(this)}
                      events={calEvents}
                      timeformat={this.props.timeformat}
                      cbNewEvent={this._newEntry.bind(this)}
                      cbDragDropEvent={this._dragDropUpdateEvent.bind(this)}
                      interactions={nextProps.interactions}
                      displayOptions={nextProps.displayOptions}
                      >
                      </EventCalendar>
        });
      });
    }
  }



  public render(): React.ReactElement<ICalendarAppProps> {

    return (
      <div>
        <div>
          {this.props.commandBarVisible?this.state.commandbar:""}
        </div>
        <div>
          <hr />
        </div>
        <div>
          {this.state.content}
        </div>
        <div>
          {this.state.panel}
        </div>
      </div>
    );
  }


  /**
   * Switches Display Mode for FullCalendar to given Type.
   * @param type
   */

  private _changeDisplayType(type: DisplayType) {
    let events = this.eventCalRef.current.getCurrentEventList();
    this.setState({
      content: <Spinner label="loading" labelPosition="bottom" size={SpinnerSize.large}></Spinner>
    });

    setTimeout(() => {
      this.setState({
        displayType: type,
        content:
          <EventCalendar
            ref={this.eventCalRef}
            displayType={type}
            cbSelectEntry={this._selectedEntry.bind(this)}
            cbUpdateEvents={this._queryEvents.bind(this)}
            events={events}
            timeformat={this.props.timeformat}
            cbNewEvent={this._newEntry.bind(this)}
            cbDragDropEvent={this._dragDropUpdateEvent.bind(this)}
            interactions={this.props.interactions}
            displayOptions={this.props.displayOptions}
          ></EventCalendar>
      }
      );
      //this.eventCalRef = React.createRef();
    }, 300
    );
  }

  /**
   * Activates the Event Panel in edit mode to display Input Form
   */
  private _newEntry(newDate:string) {
    if(!this. props.interactions.dateClickNew && typeof newDate=='string'){
      return;
    }else if(typeof newDate!='string'){
      newDate=null;
    }
    this.setState({
      panel: <EventPanel
        cbRefreshGrid={this._updateGrid.bind(this)}
        cbDelete={this._deleteEvent.bind(this)}
        cbSave={this._saveChanges.bind(this)}
        categories={this.state.categories}
        createNew={true}
        context={this.props.context}
        relativeLibOrListUrl={this.props.relativeLibOrListUrl}
        remoteSiteUrl={this.props.remoteSiteUrl}
        timeformat={this.props.timeformat}
        newDateStr={newDate}
      ></EventPanel>
    });
  }

  /**
   * Callback Method called by EventPanel to delete an Event from SharePoint
   * @param event
   */
  private _deleteEvent(event: ISPEvent): Promise<boolean> {
    let con = new PnPListConnector(this.props.listName, this.props.context, this.props.remoteSiteUrl);
    return con.deleteItem(event).then((result) => {
      return Promise.resolve(true);
    }).catch((error) => {
      console.error(error);
      return Promise.reject(false);
    });
  }

  private _dragDropUpdateEvent(event:IFullCalendarEvent){
    let spEvent:ISPEvent = EventConverter.getSPEvent(event);
    this._saveChanges(spEvent).then((result)=>{
      setTimeout(() => { }, 50);
      this._updateGrid();
    });


  }


  /**
   * Callback Method called by EventPanel to save Data to SharePoint
   * @param event
   */
  private _saveChanges(event: ISPEvent): Promise<ItemUpdateResult | ItemAddResult> {
    let con = new PnPListConnector(this.props.listName, this.props.context, this.props.remoteSiteUrl);
    let spEvent = EventConverter.getCustomEvent(event,this.props.fieldMapping);
    if (!spEvent.Id) {
      return con.addIem(spEvent).then((result)=>{
        return result;
      }).catch((error)=>{
        console.log(error);
        return error;
      });
    } else {
      return con.updateItem(spEvent.Id, spEvent).then((result)=>{
        return result;
      }).catch((error)=>{
        console.log(error);
        return error;
      });
    }

  }

  private _updateGrid() {
  this._queryEvents(this.state.displayType, this.eventCalRef.current.getDisplayDate()).then((calEvents) => {

      this.setState({
        content: <EventCalendar
                    ref={this.eventCalRef}
                    cbSelectEntry={this._selectedEntry.bind(this)}
                    cbUpdateEvents={this._queryEvents.bind(this)}
                    events={calEvents}
                    timeformat={this.props.timeformat}
                    cbNewEvent={this._newEntry.bind(this)}
                    cbDragDropEvent={this._dragDropUpdateEvent.bind(this)}
                    interactions={this.props.interactions}
                    displayOptions={this.props.displayOptions}
                    ></EventCalendar>
      });

    }).catch((error) => {
      console.error(error);
    });
  }


  private _queryEvents(displayType: DisplayType, currentDisplayDate: Date, nextProps?: ICalendarAppProps): Promise<any> {
    return this._queryMainEventList(displayType,currentDisplayDate,nextProps).then((calEvents)=>{
      this._queryOverlayDataSources(displayType,currentDisplayDate,nextProps);
      return Promise.resolve(calEvents);
    }).catch(error=>{
      return Promise.reject(error);
    })
  }

  private _queryOverlayDataSources(displayType: DisplayType, currentDisplayDate: Date, nextProps?: ICalendarAppProps){
    if(this.props.overlayListCollection ==null ){
      return;
    }

    setTimeout(()=>{
      this.props.overlayListCollection.forEach((listConfig,index)=>{
        let con = new PnPListConnector(listConfig.SiteUrl, null, listConfig.ListName);

        this._queryEventList(con,listConfig.ListName,displayType,currentDisplayDate,false).then((overlayEvents)=>{
          var eventSource:any = {
            id:"0XDATAX00",
            events: [
              {
                title: 'Event1',
                start: '2020-11-04',
                extendedProps: {
                  location: "loc1",
                  description: "desc1",
                  category: "desc1",
                  isEditable:false
                }
              },
              {
                title: 'Event2',
                start: '2020-11-05',
                extendedProps: {
                  location: "loc2",
                  description: "desc2",
                  category: "desc2",
                  isEditable:false
                }
              }
            ],
            color: listConfig.BackgroundColor,   // an option!
            textColor: listConfig.TextColor // an option!
          }
          this.eventCalRef.current.removeEventSourceById("0XDATAX00");
          this.eventCalRef.current.addEventSource(eventSource);
        }).catch(error=>{
          console.error("Cannot Query Overlay Events");
        });
      });
    },2000);
  }


  private _queryMainEventList(displayType: DisplayType, currentDisplayDate: Date, nextProps?: ICalendarAppProps): Promise<any> {
    let listName = nextProps ? nextProps.listName : this.props.listName;
    let relativeLibOrListUrl = nextProps ? nextProps.relativeLibOrListUrl : this.props.relativeLibOrListUrl;
    let remoteSiteUrl = nextProps ? nextProps.remoteSiteUrl : this.props.remoteSiteUrl;
    if (!listName || !relativeLibOrListUrl) { return Promise.resolve([]); }
    let con = new PnPListConnector(this.props.listName, this.props.context, this.props.remoteSiteUrl);
    let siteCon: SiteConnector = new SiteConnector(this.props.context);

    siteCon.getColumnOptions(this.props.fieldMapping.Category, listName, remoteSiteUrl).then((categories) => {
      let categoryValues = categories.value[0].Choices.map((item) => {
        return {
          key: item,
          text: item
        };
      });
      this.setState({ categories: categoryValues });
    });
/*
    let startDate;
    let endDate;
    switch (+displayType) {
      case DisplayType.WeekGrid:
        startDate = moment(currentDisplayDate).startOf('month').format("YYYY-MM-DD");
        endDate = moment(currentDisplayDate).endOf('month').format("YYYY-MM-DD");
        break;
      default:
        startDate = moment(currentDisplayDate).startOf('week').format("YYYY-MM-DD");
        endDate = moment(currentDisplayDate).endOf('week').format("YYYY-MM-DD");
        break;
    }
   var camlBuilder = new CamlBuilder();
   var caml: string = camlBuilder.Where()
   .DateField(this.props.fieldMapping["EventDate"]).GreaterThan(moment(startDate).toDate())
   .And()
   .DateField(this.props.fieldMapping["EndDate"]).LessThanOrEqualTo(moment(endDate).toDate()).ToString();
    caml = `<View>${caml}</View>`;
    return con.getItemByCAML(listName, { ViewXml: caml }).then((result) => {
      let calEvents = result.map((event) => {
        return EventConverter.getFCEvent(event,this.props.fieldMapping);
      });
      return Promise.resolve(calEvents);
    }).catch((error) => {
      Promise.reject(error);
    });
    */
   return this._queryEventList(con,listName,displayType,currentDisplayDate,true);
  }

  private _queryEventList(con:PnPListConnector,listName:string,displayType: DisplayType, currentDisplayDate: Date,eventsEditable:boolean): Promise<any>{
    let startDate;
    let endDate;
    switch (+displayType) {
      case DisplayType.WeekGrid:
        startDate = moment(currentDisplayDate).startOf('month').format("YYYY-MM-DD");
        endDate = moment(currentDisplayDate).endOf('month').format("YYYY-MM-DD");
        break;
      default:
        startDate = moment(currentDisplayDate).startOf('week').format("YYYY-MM-DD");
        endDate = moment(currentDisplayDate).endOf('week').format("YYYY-MM-DD");
        break;
    }
   var camlBuilder = new CamlBuilder();
   var caml: string = camlBuilder.Where()
   .DateField(this.props.fieldMapping["EventDate"]).GreaterThan(moment(startDate).toDate())
   .And()
   .DateField(this.props.fieldMapping["EndDate"]).LessThanOrEqualTo(moment(endDate).toDate()).ToString();
    caml = `<View>${caml}</View>`;
    return con.getItemByCAML(listName, { ViewXml: caml }).then((result) => {
      let calEvents = result.map((event) => {
        return EventConverter.getFCEvent(event,this.props.fieldMapping,eventsEditable);
      });
      return Promise.resolve(calEvents);
    }).catch((error) => {
      Promise.reject(error);
    });
  }


  private _selectedEntry(entry: IFullCalendarEvent) {
    this.setState({
      panel: <EventPanel
        cbRefreshGrid={this._updateGrid.bind(this)}
        cbDelete={this._deleteEvent.bind(this)}
        cbSave={this._saveChanges.bind(this)}
        categories={this.state.categories}
        event={entry}
        context={this.props.context}
        relativeLibOrListUrl={this.props.relativeLibOrListUrl}
        remoteSiteUrl={this.props.remoteSiteUrl}
        timeformat={this.props.timeformat}
      ></EventPanel>
    });

  }






}
