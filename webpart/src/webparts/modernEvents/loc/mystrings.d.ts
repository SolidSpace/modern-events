declare interface IModernEventsWebPartStrings {
  DescriptionFieldLabel: string;
  PropertyPaneDescription: string;
  DescriptionFieldLabel: string;
  PPaneListPage: string;
  PPaneDisplayOptionsPage: string;
  //Configure Texts and Labels
  LabelConfigIconName: string;
  LabelConfigIconText: string;
  LabelConfigIconName: string;
  LabelConfigIconDescription: string;
  LabelConfigBtnLabel: string;
  // SiteGroups
  SiteGroupName: string;
  SiteGroupDataBinding: string,
  SiteGroupCalDisplayOptions: string,
  CommandbarGroupName: string;
  DisplayGroupName: string;
  InteractionGroupName: string;
  // Labels
  LabelViewButtons: string;
  LabelTimeformat: string;
  LabelWeekStart: string;
  LabelSite: string;
  LabelSiteOther: string;
  LabelListTitle: string;
  LabelSite: string;
  LabelCommandbar: string;
  LabelViewMonth: string;
  LabelViewWeek: string;
  LabelViewList: string;
  LabelInterActionEventClickNew: string;
  LabelInterActionEventDragDrop: string;
  LabelCustListFieldMap:string;
  LabelUseCustomList: string;
  LabelCustListTitle:string;
  LabelCustListCategory:string;
  LabelCustListLocation:string;
  LabelCustListStart:string;
  LabelCustListEnd:string;
  LabelCustListAllDayEvent:string;
  LabelCustListDescription:string;
  //Commandbar Labels
  LabelButtonNew: string;
  LabelButtonNewEvent: string;
  LabelButtonMonth: string;
  LabelButtonTime: string;
  LabelButtonList: string;
  //EventPanel Labels
  LabelEventTitle: string;
  LabelEventCategory: string;
  LabelEventLocation: string;
  LabelEventAllDay: string;
  LabelEventStartDate: string;
  LabelEventEndDate: string;
  LabelEventDescription: string;
  //Event Panel ComboBox Values
  WeekDay0:string;
  WeekDay1:string;
  WeekDay2:string;
  WeekDay3:string;
  WeekDay4:string;
  WeekDay5:string;
  WeekDay6:string;
  WeekDay7:string;
  //Tooltips
  TooltipCancel: string;
  TooltipEdit: string;
  TooltipSave: string;
  TooltipDelete: string;
  //Fullcalendar
  LabelAllDay: string;
  //Errors
  ErrIENotCompatible:string;
}

declare module 'ModernEventsWebPartStrings' {
  const strings: IModernEventsWebPartStrings;
  export = strings;
}
