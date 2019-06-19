import * as React from 'react';
import * as ReactDom from 'react-dom';
import { escape } from '@microsoft/sp-lodash-subset';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, IPropertyPaneDropdownOption, PropertyPaneDropdown, PropertyPaneCheckbox, PropertyPaneLabel } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { Placeholder, IPlaceholderProps } from "@pnp/spfx-controls-react/lib/Placeholder";
import * as strings from 'ModernEventsWebPartStrings';
//import { element } from 'prop-types';
import { CalendarApp, ICalendarAppProps } from './components/CalendarApp';
import "./sass/style.scss";
import { DisplayType } from './components/ENUMDisplayType';
//import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { SiteConnector, ISPList } from './services/SiteConnector';
import { string } from 'prop-types';
//import { RectangleEdge } from 'office-ui-fabric-react/lib/utilities/positioning';

export interface IModernEventsWebPartProps {
  site: string;
  siteOther: string;
  listTitle: string;
  description: string;
  commandbar: boolean;
  viewMonth: boolean;
  viewWeek: boolean;
  viewList: boolean;
  timeformat:string;
  interactionEventClick:boolean
}
export default class ModernEventsWebPart extends BaseClientSideWebPart<IModernEventsWebPartProps> {
  private _siteOptions: IPropertyPaneDropdownOption[] = [];
  private _listOptions: IPropertyPaneDropdownOption[] = [];
  private _listDisabled = true;
  private _otherDisabled = true;

  public render(): void {
    if(this.properties.site && this.properties.listTitle){
      const app: React.ReactElement<ICalendarAppProps> = React.createElement(
        CalendarApp,
        {
          context: this.context,
          remoteSiteUrl: this.properties.site,
          relativeLibOrListUrl: "/lists/" + this.properties.listTitle,
          displayType: DisplayType.WeekGrid,
          listName: this.properties.listTitle,
          timeformat:this.properties.timeformat,
          commandBarVisible:this.properties.commandbar,
          commandBarButtonVisibility:{
            month:this.properties.viewMonth,
            time:this.properties.viewWeek,
            list:this.properties.viewList
          },
          interactions:{dateClickNew:this.properties.interactionEventClick?this.properties.interactionEventClick:true}
        }
      );
      ReactDom.render(app, this.domElement);
    } else{
      const configure:React.ReactElement<IPlaceholderProps> = React.createElement(
        Placeholder,{
          iconName:strings.LabelConfigIconName,
          iconText:strings.LabelConfigIconText,
          description:strings.LabelConfigIconDescription,
          buttonLabel:strings.LabelConfigBtnLabel,
          onConfigure:this._onConfigureWebpart.bind(this)
        }
      );
      ReactDom.render(configure, this.domElement);
    }
  }

  private _onConfigureWebpart(){
    this.context.propertyPane.open();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected onPropertyPaneConfigurationStart(): void {
    let con: SiteConnector = new SiteConnector(this.context);
    con.getSiteRootWeb().then((rootweb) => {
      con.getSites(rootweb['Url']).then((sitesResult) => {
        var sites: IPropertyPaneDropdownOption[] = [];
        sites.push({ key: this.context.pageContext.web.absoluteUrl, text: 'This Site' });
       // sites.push({ key: 'other', text: 'Other Site (Specify Url)' });
        for (var _key in sitesResult.value) {
          if (this.context.pageContext.web.absoluteUrl != sitesResult.value[_key]['Url']) {
            sites.push({ key: sitesResult.value[_key]['Url'], text: sitesResult.value[_key]['Title'] });
          }
        }
        this._siteOptions = sites;
        this.context.propertyPane.refresh();
        //let siteUrl = this.properties.site;
        if (this.properties.site) {
          con.getCalendarListTitles(this.properties.site).then((listTitleResult) => {
            this._listOptions = listTitleResult.value.map((list: ISPList) => {
              return {
                key: list.Title,
                text: list.Title
              };
            });
            this.context.propertyPane.refresh();
            this.render();
          });
        } else {
          this.context.propertyPane.refresh();

          this.render();
        }
      });
    });
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {

    if (newValue == 'other') {
      this._otherDisabled = false;
      this.properties.listTitle = null;
    } else if (oldValue === 'other' && newValue != 'other') {
      this._otherDisabled = true;
      this.properties.siteOther = null;
      this.properties.listTitle = null;
    }
    let con: SiteConnector = new SiteConnector(this.context);
    if ((propertyPath === 'site' || propertyPath === 'other') && newValue) {
      this._listDisabled = true;
      var siteUrl = newValue;
      con.getCalendarListTitles(this.properties.site).then((listTitleResult) => {
        this._listOptions = listTitleResult.value.map((list: ISPList) => {
          return {
            key: list.Title,
            text: list.Title
          };
        });

        this._listDisabled = false;
        this.context.propertyPane.refresh();
        this.render();
      });
    }


  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }



  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.SiteGroupName,
              groupFields: [
                PropertyPaneDropdown('site', {
                  label: strings.LabelSite,
                  options: this._siteOptions
                }),
                /*
                PropertyPaneTextField('siteOther', {
                  label: strings.LabelSiteOther,
                  ariaLabel: "otherSiteAria",
                  disabled: this._otherDisabled

                }),*/
                PropertyPaneDropdown('listTitle', {
                  label: strings.LabelListTitle,
                  options: this._listOptions,
                  disabled: this._listDisabled
                }),
              ]
            },
            {
              groupName: strings.DisplayGroupName,
              groupFields: [
                PropertyPaneDropdown('timeformat', {
                  selectedKey:"24h",
                  label: strings.LabelTimeformat,
                  options: [{key:'24h',text:'24 Hours'},{key:'12h',text:'12 Hours AM/PM'}],
                  disabled: false
                })
              ]
            },
            {
              groupName: strings.InteractionGroupName,
              groupFields: [
                PropertyPaneCheckbox('interactionEventClick', {
                  text: strings.LabelInterActionEventClickNew,
                  checked: true,
                  disabled: false
                }),
              ]
            },
            {
              groupName: strings.CommandbarGroupName,
              groupFields: [
                PropertyPaneCheckbox('commandbar', {
                  text: strings.LabelCommandbar,
                  checked: false
                }),
                PropertyPaneLabel('viewMonth',{
                  text:strings.LabelViewButtons
                }),
                PropertyPaneCheckbox('viewMonth', {
                  text: strings.LabelViewMonth,
                  checked: false,
                  disabled: !this.properties.commandbar
                }),
                PropertyPaneCheckbox('viewWeek', {
                  text: strings.LabelViewWeek,
                  checked: false,
                  disabled: !this.properties.commandbar
                }),
                PropertyPaneCheckbox('viewList', {
                  text: strings.LabelViewList,
                  checked: false,
                  disabled: !this.properties.commandbar
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
