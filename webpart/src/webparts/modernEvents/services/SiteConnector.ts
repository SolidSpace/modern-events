import {SPHttpClient, SPHttpClientResponse} from '@microsoft/sp-http';

export interface ISPLists {
  value: ISPList[];
}
export interface ISPList {
  Title: string;
  Id: string;
}

export class SiteConnector{

  private context:any=null;

  constructor(context:any){
    this.context=context;
  }

  public getSiteRootWeb(): Promise<ISPLists> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/Site/RootWeb?$select=Title,Url`, SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) => {
      return response.json();
    });
  }

  public getSites(rootWebUrl: string): Promise<ISPLists> {
    return this.context.spHttpClient.get(rootWebUrl + `/_api/web/webs?$select=Title,Url`, SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) => {
      return response.json();
    });
  }

  public getCalendarListTitles(site: string): Promise<ISPLists> {
    return this.context.spHttpClient.get(site + `/_api/web/lists?$filter=Hidden eq false and BaseType eq 0 and BaseTemplate eq 106`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  /**
   * Query All Lists from a Site
   * @param site
   */
  public getListTitles(site: string): Promise<ISPLists> {
    return this.context.spHttpClient.get(site + `/_api/web/lists?$filter=Hidden eq false and BaseType eq 0`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  public getListTitlesByTemplate(site: string,templateId:string): Promise<ISPLists> {
    return this.context.spHttpClient.get(site + `/_api/web/lists?select=Title,ServerRelativeUrl&$filter=Hidden eq false and BaseTemplate eq `+templateId, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  public getListFormProperties(site: string,listName:string): Promise<string>{
    return this.context.spHttpClient.get(site+"/_api/web/lists/GetByTitle('" + listName + "')/Forms?$select=ServerRelativeUrl", SPHttpClient.configurations.v1)
    .then((response:SPHttpClientResponse)=>{
      return response.json();
    });
    //"/_api/web/lists/GetByTitle('" + listName + "')/Forms?$select=ServerRelativeUrl"
  }

//"BaseTemplate": 106 = Calendar List
//"BaseTemplate": 100 = Custom List


  /**
   * Query All Columns from a List
   * @param listNameColumns
   * @param listsite
   */
  public getListColumns(listNameColumns: string,listsite: string): Promise<any> {
    return this.context.spHttpClient.get(listsite + `/_api/web/lists/GetByTitle('${listNameColumns}')/Fields?$filter=Hidden eq false and ReadOnlyField eq false`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  /**
   * Query following List Columns necessary for Event Mapping
   * FieldTypeKind	:	2 = Text
   * FieldTypeKind	:	4 = DateTime
   * FieldTypeKind	:	3 = Multiline
   * @param listNameColumns
   * @param listsite
   */
  public getEventListColumns(listNameColumns: string,listsite: string): Promise<any> {
    return this.context.spHttpClient.get(listsite + `/_api/web/lists/GetByTitle('${listNameColumns}')/Fields?$filter=Hidden eq false and ReadOnlyField eq false and (FieldTypeKind eq 2 or FieldTypeKind eq 3 or FieldTypeKind eq 4 or FieldTypeKind eq 6 or FieldTypeKind eq 8)`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  /**
   *
   * @param columnName
   * @param listNameColumns
   * @param listsite
   */
  public getColumnOptions(columnName:string,listNameColumns: string,listsite: string): Promise<any> {
    return this.context.spHttpClient.get(listsite + `/_api/web/lists/GetByTitle('${listNameColumns}')/Fields?$filter=EntityPropertyName eq '${columnName}'`, SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) => {
      return response.json();
    });
  }
}

