import * as React from 'react';
import * as moment from 'moment';
import * as strings from 'ModernEventsWebPartStrings';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { IFullCalendarEvent } from "./IFullCalendarEvent";
import { ISPEvent } from "./ISPEvent";
import { DateTimePicker, DateConvention, TimeConvention } from '@pnp/spfx-controls-react/lib/dateTimePicker';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { DefaultButton, PrimaryButton, IButtonProps, Button } from 'office-ui-fabric-react/lib/Button';
import { SecurityTrimmedControl, PermissionLevel } from "@pnp/spfx-controls-react/lib/SecurityTrimmedControl";
import { SPPermission } from '@microsoft/sp-page-context';
import CKEditor from 'ckeditor4-react';
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { EventConverter } from "../util/EventConverter";
import { Icon } from 'office-ui-fabric-react/lib/Icon';

export interface IEventPanelProps {
  cbSave: any;
  cbDelete: any;
  cbRefreshGrid: any;
  categories: IDropdownOption[];
  context: any;
  remoteSiteUrl: string;
  relativeLibOrListUrl: string;
  event?: IFullCalendarEvent;
  createNew?: boolean;
  timeformat: string;
}

export interface IEventPanelState {
  isEditMode: boolean;
  isOpen: boolean;
  isDialogOpen: boolean;
  isSaveInProgress:boolean;
  event: ISPEvent;
  inputValidation: boolean;
}

export class EventPanel extends React.Component<IEventPanelProps, IEventPanelState> {

  constructor(props: IEventPanelProps) {
    super(props);
    if (this.props.event != null) {
      this.state = {
        isEditMode: false,
        isOpen: true,
        isDialogOpen: false,
        event: EventConverter.getSPEvent(props.event),
        inputValidation: false,
        isSaveInProgress:false
      };
    } else {
      this.state = {
        isEditMode: true,
        isOpen: true,
        isDialogOpen: false,
        event: {
          EventDate: moment().format("YYYY-MM-DDTHH:mm:ss"),
          EndDate: moment().add(1, 'hour').format("YYYY-MM-DDTHH:mm:ss"),
          Title: "New Event",
          fAllDayEvent: false,
          Location: "",
          Category: "",
          Description: ""
        },
        inputValidation: false,
        isSaveInProgress:false
      };
    }
  }

  /**
   * Validates CreateNew Value to determine if the user wants to create a new
   * Event or has opend an exisiting Event
   * @param nextProps
   */
  public componentWillReceiveProps(nextProps: IEventPanelProps) {
    if (nextProps.createNew) {
      this.setState({
        isEditMode: true,
        isOpen: true,
        event: {
          EventDate: moment().format("YYYY-MM-DDTHH:mm:ss"),
          EndDate: moment().add(1, 'hour').format("YYYY-MM-DDTHH:mm:ss"),
          Title: "New Event",
          fAllDayEvent: false,
          Location: "",
          Category: "",
          Description: ""
        },
        inputValidation: false
      });
    } else {

      let newEvent = (nextProps.event && nextProps.event != this.props.event) ? EventConverter.getSPEvent(nextProps.event) : EventConverter.getSPEvent(this.props.event);

      this.setState({
        ...this.state,
        event: newEvent,
        isOpen: true,
        inputValidation: false
      });
    }

  }



  public render(): React.ReactElement<IEventPanelProps> {
    let timeConvention = this.props.timeformat == '24h' ? TimeConvention.Hours24 : TimeConvention.Hours12;
    return (
      <Panel isOpen={this.state.isOpen} onAbort={() => console.log("ABORTING PANEL!")}>
        <div className="ms-Grid" dir="ltr">
          <div className="ms-Grid-row">
            <div className="ms-Grid-col">
              <h3>
                {this.state.event.Title}
              </h3>
            </div>
          </div>
          <div className="ms-Grid-row">
            <div className="ms-Grid-col overwriteGridShrinking">
              <TextField label={strings.LabelEventTitle}
                disabled={!this.state.isEditMode}
                defaultValue={this.state.event.Title}
                onChanged={this._onChangeTitle.bind(this)}
                style={(this.state.inputValidation && (this.state.event.Title.trim() == '')) ? { backgroundColor: '#FFE2E7' } : {}}
              />
            </div>
          </div>
          <div className="ms-Grid-row">
            <div className="ms-Grid-col overwriteGridShrinking">
              <Dropdown
                label={strings.LabelEventCategory}
                defaultSelectedKey={this.state.event.Category}
                options={this.props.categories}
                disabled={!this.state.isEditMode}
                onChange={this._onChangeCategory.bind(this)}
              />
            </div>
          </div>

          <div className="ms-Grid-row">
            <div className="ms-Grid-col overwriteGridShrinking">
              <TextField label={strings.LabelEventLocation}
                disabled={!this.state.isEditMode}
                defaultValue={this.state.event.Location}
                onChanged={this._onChangeLocation.bind(this)}
              />
            </div>
          </div>
          <div className="ms-Grid-row">
            <div className="ms-Grid-col">
              <hr />
            </div>
          </div>
          <div className="ms-Grid-row">
            <div className="ms-Grid-col">
              <Checkbox
                label={strings.LabelEventAllDay}
                checked={this.state.event.fAllDayEvent}
                onChange={this._onControlledCheckboxChange.bind(this)}
                name="cbxAllDayEvent"
                disabled={!this.state.isEditMode}
              />
            </div>
          </div>
          <div className="ms-Grid-row">
            <div className="ms-Grid-col">
              <DateTimePicker label={strings.LabelEventStartDate}
                dateConvention={this.state.event.fAllDayEvent ? DateConvention.Date : DateConvention.DateTime}
                timeConvention={timeConvention}
                value={moment(this.state.event.EventDate).toDate()}
                disabled={!this.state.isEditMode}
                onChange={this._onChangeStartDate.bind(this)}
                showGoToToday={true}

              />
            </div>
          </div>
          <div className="ms-Grid-row">
            <div className="ms-Grid-col">
              <DateTimePicker label={strings.LabelEventEndDate}
                dateConvention={this.state.event.fAllDayEvent ? DateConvention.Date : DateConvention.DateTime}
                timeConvention={timeConvention}
                value={moment(this.state.event.EndDate).toDate()}
                disabled={!this.state.isEditMode}
                onChange={this._onChangeEndDate.bind(this)}
              />
            </div>
          </div>
          <div className="ms-Grid-row">
            <div className="ms-Grid-col">
              <hr />
            </div>
          </div>
          <div className="ms-Grid-row">
            <div className="ms-Grid-col">

              <Label>{strings.LabelEventDescription}</Label>
              <CKEditor
                data={this.state.event.Description}
                config={{
                  toolbar: [['Bold', 'Italic', 'Underline', 'Strike', 'Subscript', 'Superscript', '-', 'CopyFormatting', 'RemoveFormat']]
                }}
                readOnly={!this.state.isEditMode}
                onChange={this._onFCKChange.bind(this)}
              />
            </div>
          </div>
          <div className="ms-Grid-row">
            <div className="ms-Grid-col">
              {this.state.isDialogOpen ? this._getDialog() : this._getActionButtons()}
            </div>
          </div>
        </div>
      </Panel >
    );
  }
  private _getDialog(): JSX.Element {
    return (
      <div>
        <div className="se-PanelAction">
          <PrimaryButton
            data-automation-id="delete"
            allowDisabledFocus={true}
            disabled={false}
            checked={true}
            text="Delete"
            onClick={this._deleteEvent.bind(this)}
            hidden={true}
          />
        </div>
        <div className="se-PanelAction">
          <Button
            data-automation-id="Close"
            allowDisabledFocus={true}
            disabled={false}
            checked={true}
            text="Cancel"
            onClick={() => this.setState({ isDialogOpen: false })}
          />
        </div>
      </div>
    );
  }
  /**
   * Returns current Action Buttons depending on EditMode State
   */
  private _getActionButtons(): JSX.Element {
    if(this.state.isSaveInProgress){
      return <div className="se-PanelActions"><Spinner label="save in progress" labelPosition="bottom" size={SpinnerSize.large}></Spinner></div>;
    }
    let relativeSiteUrl = this.props.remoteSiteUrl ? this.props.remoteSiteUrl.replace(/https:\/\/.+.sharepoint.com/g, "") + this.props.relativeLibOrListUrl : "";
    relativeSiteUrl = !relativeSiteUrl.substr(1, 1).match("/") ? relativeSiteUrl : "/" + relativeSiteUrl;
    if (this.state.isEditMode) {
      return (
        <div className="se-PanelActions">
          <div className="se-PanelAction ">
            <div className="tooltip">
              <Icon iconName="Accept" className="se-PanelAction-Primary ms-font-xxl" onClick={this._save.bind(this)} />
              <span className="tooltiptext">{strings.TooltipSave}</span>
            </div>
          </div>
          <div className="se-PanelAction ">
            <div className="tooltip">
              <Icon iconName="Cancel" className="se-PanelAction-Secondary ms-font-xxl" onClick={this._closePanel.bind(this)} />
              <span className="tooltiptext">{strings.TooltipCancel}</span>
            </div>
          </div>
        </div>
      );
    } else {
      return (
        <div className="se-PanelActions">
          <div className="se-PanelAction ">
            <SecurityTrimmedControl context={this.props.context}
              level={PermissionLevel.remoteListOrLib}
              remoteSiteUrl={this.props.remoteSiteUrl}
              relativeLibOrListUrl={relativeSiteUrl}
              permissions={[SPPermission.addListItems]}>
              <div className="tooltip">
                <Icon iconName="WindowEdit" className="se-PanelAction-Primary ms-font-xxl" onClick={this._toggleEdit.bind(this)} />
                <span className="tooltiptext">{strings.TooltipEdit}</span>
              </div>
            </SecurityTrimmedControl>
          </div>
          <div className="se-PanelAction ">
            <SecurityTrimmedControl context={this.props.context}
              level={PermissionLevel.remoteListOrLib}
              remoteSiteUrl={this.props.remoteSiteUrl}
              relativeLibOrListUrl={relativeSiteUrl}
              permissions={[SPPermission.addListItems]}>
              <div className="tooltip">
                <Icon iconName="Delete" className="se-PanelAction-Primary ms-font-xxl" onClick={this._deleteEvent.bind(this)} />
                <span className="tooltiptext">{strings.TooltipDelete}</span>
              </div>
            </SecurityTrimmedControl>
          </div>
          <div className="se-PanelAction ">
            <div className="tooltip">
              <Icon iconName="Cancel" className="se-PanelAction-Secondary ms-font-xxl" onClick={this._closePanel.bind(this)} />
              <span className="tooltiptext">{strings.TooltipCancel}</span>
            </div>
          </div>
        </div>


      );
    }
  }




  /**
   * OnClick Event for save Button is executed
   * Calls the Callback Function provided from App Class
   */
  private _save(): void {
    if (!this.state.event.Title || this.state.event.Title.trim() == '') {
      this.setState({ inputValidation: true });
      return;
    }
    this.setState({isSaveInProgress:true});

    this.props.cbSave(this.state.event).then((result) => {
      setTimeout(() => { }, 50);
      this.props.cbRefreshGrid();
      this.setState({
        ...this.state,
        isEditMode: !this.state.isEditMode,
        isSaveInProgress:false
      });

    });
  }

  /**
   * OnClick Event for edit Buton is executed
   * Switches the current Event into Edit mode
   * @param event
   */
  private _toggleEdit(event: any): void {
    this.setState({
      ...this.state,
      isEditMode: !this.state.isEditMode
    });
  }

  /**
   * Displays a dialog to to ask user before deleting the event.
   */
  private _deleteEvent() {
    this.props.cbDelete(this.state.event).then((result) => {
      setTimeout(() => { }, 50);
      this.props.cbRefreshGrid();
      this.setState({ isDialogOpen: false, isOpen: false });
    });
  }


  /** On Change Event Handler Section */

  /**
   * OnChange Event fired while editing Title Field
   * @param newValue
   */
  private _onChangeTitle(newValue: any) {
    this.setState({
      event: {
        ...this.state.event,
        Title: newValue
      }
    });
  }

  /**
   * OnChange Event fired when a new Category is selected
   * @param event
   * @param item
   */
  private _onChangeCategory(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption) {
    this.setState({
      event: {
        ...this.state.event,
        Category: item.key.toString()
      }
    });

  }

  /**
   * OnChange Event fired when StartDate is changed
   * @param dateTimeValue
   */
  private _onChangeStartDate(dateTimeValue: any) {
    let current = moment(this.state.event.EventDate);
    let start = moment(dateTimeValue);
    let end = moment(this.state.event.EndDate);
    if (start.isSame(current)) { return; }
    this.setState({
      event: {
        ...this.state.event,
        EventDate: start.format("YYYY-MM-DDTHH:mm:ss"),
        EndDate: (start.isAfter(end)) ? start.format("YYYY-MM-DDTHH:mm:ss") : end.format("YYYY-MM-DDTHH:mm:ss")
      }
    });
  }

  /**
  * OnChange Event fired when EndDate is changed
  * @param dateTimeValue
  */
  private _onChangeEndDate(dateTimeValue: any) {
    let current = moment(this.state.event.EndDate);
    let start = moment(this.state.event.EventDate);
    let end = moment(dateTimeValue);
    if (end.isSame(current)) { return; }
    this.setState({
      event: {
        ...this.state.event,
        EventDate: start.format("YYYY-MM-DDTHH:mm:ss"),
        EndDate: (start.isAfter(end)) ? start.format("YYYY-MM-DDTHH:mm:ss") : end.format("YYYY-MM-DDTHH:mm:ss")
      }
    });
  }



  /**
   * OnChange handler fired when changing the value of Location Field
   * @param newValue
   */
  private _onChangeLocation(newValue: any) {
    this.setState({
      event: {
        ...this.state.event,
        Location: newValue
      }
    });
  }

  /**
   * OnClick Event for Close Button
   */
  private _closePanel() {
    this.setState({
      ...this.state,
      event: {
        EndDate: "",
        EventDate: "",
        Title: "",
        fAllDayEvent: false,
        Location: "",
        Description: "",
        Category: ""
      },
      isOpen: false,
      isEditMode: false
    });
  }

  /**
   * Change Event when someon edits the Richtext FCK Editor Value
   * @param evt
   * @param editor
   */
  private _onFCKChange(evt: any, editor: any) {
    this.setState({
      event: {
        ...this.state.event,
        Description: evt.editor.getData()
      }
    });

  }

  /**
   * Change Event fired when AllDay Checkbox Value is changed
   * @param evt
   * @param checked
   */
  private _onControlledCheckboxChange(evt: React.FormEvent<HTMLElement>, checked: boolean) {

    if (checked) {
      this.setState({
        event: {
          ...this.state.event,
          fAllDayEvent: checked,
          EventDate: moment(this.state.event.EventDate).set({ h: 0, m: 0 }).format("YYYY-MM-DDTHH:mm:ss"),
          EndDate: moment(this.state.event.EndDate).set({ h: 23, m: 59 }).format("YYYY-MM-DDTHH:mm:ss")
        }
      });
    } else {
      this.setState({
        event: {
          ...this.state.event,
          fAllDayEvent: checked
        }
      });
    }

  }
}
