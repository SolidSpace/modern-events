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

  public getListTitles(site: string): Promise<ISPLists> {
    return this.context.spHttpClient.get(site + `/_api/web/lists?$filter=Hidden eq false and BaseType eq 0`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  public getListColumns(listNameColumns: string,listsite: string): Promise<any> {
    return this.context.spHttpClient.get(listsite + `/_api/web/lists/GetByTitle('${listNameColumns}')/Fields?$filter=Hidden eq false and ReadOnlyField eq false`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  public getColumnOptions(columnName:string,listNameColumns: string,listsite: string): Promise<any> {
    return this.context.spHttpClient.get(listsite + `/_api/web/lists/GetByTitle('${listNameColumns}')/Fields?$filter=EntityPropertyName eq '${columnName}'`, SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) => {
      return response.json();
    });
  }
}

