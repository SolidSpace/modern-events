import * as React from 'react';
import { DisplayType } from './ENUMDisplayType';
import { CommandBar } from 'office-ui-fabric-react/lib/CommandBar';
import * as strings from 'ModernEventsWebPartStrings';
import { ICBButtonVisibility } from "./ICBButtonVisibility";

export interface ICalendarCommandbarProps {
  cbNewEntry: any;
  cbWeekGrid: any;
  cbTimeGrid: any;
  cbListGrid: any;
  buttonVisibiliy?: ICBButtonVisibility;
}

export interface ICalendarCommandbarState {
  buttonVisibiliy: ICBButtonVisibility;
}

export class CalendarCommandbar extends React.Component<ICalendarCommandbarProps, ICalendarCommandbarState> {

  constructor(props: ICalendarCommandbarProps) {
    super(props);
    this.state = {
      buttonVisibiliy: {
        list: this.props.buttonVisibiliy.list || false,
        month: this.props.buttonVisibiliy.month || false,
        time: this.props.buttonVisibiliy.time || false
      }
    };
  }

  public componentWillReceiveProps(nextProps: ICalendarCommandbarProps) {
    this.setState({
      buttonVisibiliy: {
        list: nextProps.buttonVisibiliy.list ,
        month: nextProps.buttonVisibiliy.month,
        time: nextProps.buttonVisibiliy.time
      }
    });
  }


  // Get Far Buttonset for CommandBar
  private _getItems = () => {
    return [
      {
        key: 'newItem',
        name: 'New',
        cacheKey: 'newitem', // changing this key will invalidate this items cache
        iconProps: {
          iconName: 'Add'
        },
        ariaLabel: 'New',
        subMenuProps: {
          items: [
            {
              key: 'calendarentry',
              name: 'Create Event',
              iconProps: {
                iconName: 'BuildQueueNew'
              },
              onClick: this.props.cbNewEntry,
              ['data-automation-id']: 'newCalEntry'
            }
          ]
        }
      }
    ];
  }

  // Get Buttonset for CommandBar
  private _getFarItems = () => {
    let buttonSet:any[]=[];
    if(this.state.buttonVisibiliy.month){
      buttonSet.push(
        {
          key: 'monthgrid',
          name: strings.LabelButtonMonth,
          ariaLabel: 'month',
          iconProps: {
            iconName: 'CalendarWeek',
          },
          onClick: () => this.props.cbWeekGrid(DisplayType.WeekGrid)
        }
      );
    }
    if(this.state.buttonVisibiliy.list){
      buttonSet.push(
        {
          key: 'timegrid',
          name: strings.LabelButtonTime,
          ariaLabel: 'time',
          iconProps: {
            iconName: 'Tiles'
          },
          iconOnly: false,
          onClick: () => this.props.cbWeekGrid(DisplayType.TimeGrid)
        }
      );
    }
    if(this.state.buttonVisibiliy.time){
      buttonSet.push(
        {
          key: 'listgrid',
          name: strings.LabelButtonList,
          ariaLabel: 'List',
          iconProps: {
            iconName: 'TimelineMatrixView'
          },
          iconOnly: false,
          onClick: () => this.props.cbWeekGrid(DisplayType.ListGrid)
        }
      );
    }
    return buttonSet;
  }
  public render(): JSX.Element {
    return (
      <div>
        <CommandBar
          items={this._getItems()}
          overflowButtonProps={{ ariaLabel: 'More commands' }}
          farItems={this._getFarItems()}
          ariaLabel={'Use left and right arrow keys to navigate between commands'}
        />
      </div>
    );
  }
}
