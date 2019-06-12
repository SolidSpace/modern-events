/*
    SharePoint API LIST Connector

    Dependencies: PNP-JS-CORE
    https://pnp.github.io/pnpjs/
    Install : npm install @pnp/pnpjs --save



    29.11.2018 - Version 2.0.0
    Migrate to pnpjs
    https://github.com/pnp/pnpjs/blob/dev/packages/pnpjs/docs/index.md
    https://pnp.github.io/pnpjs/
*/
//import pnp, { Web, List, ListEnsureResult, ItemAddResult, FieldAddResult, ItemUpdateResult, FieldCreationProperties, CamlQuery, AlreadyInBatchException, ViewAddResult } from "sp-pnp-js";
import { sp, Web, List, ItemAddResult, ItemUpdateResult, CamlQuery, ListEnsureResult, ViewAddResult } from "@pnp/sp";

import { WebPartContext } from '@microsoft/sp-webpart-base';

const LIST_EXISTS: string = 'List exists';
/**
 * Interface for List Configuration.
 * List Configuration will be used to create a List
 * in your Site Collection
 * fields - Array of Field Configurations
 * listName - Name of the List to be created
 *
 */
export interface IListCfg {
  fields: IListFieldCfg[];
  listName: string;
}

/**
 * - title - Field Name
 * - fieldType - Field Type Definition
 *               See further details https://msdn.microsoft.com/en-us/library/office/dn600182.aspx
 * - fieldProperties - Additional Field Properties
 */

export interface IListFieldCfg {
  title: string;
  fieldType: string;
  fieldProperties: & { FieldTypeKind: number; };
}

/**
 * ViewName - Name of your View
 * RowLimit - Max amout of Entries to show
 * DefaultView - Should the View be set as Default View
 * Fields -
 */
export interface IViewConfig {
  ViewName: string;
  RowLimit: number;
  DefaultView: boolean;
  ViewQuery: string;
  Fields: string[];
}

/**
 * List Configuration Class
 */
class BaseListConfig implements IListCfg {
  public fields: IListFieldCfg[];
  public listName: string;
}

/**
 * Class to connect to the SharePoint API and manage your lists and Items
 */
export class PnPListConnector {
  private properties: IListCfg;
  private createOnFail: boolean;
  private web: Web = sp.web;

  /**
   * Supply either a Webpart Context to use PNPListConnector with your current
   * Site Collection or a path to the remote site url.
   * @param context
   * @param siteUrl
   */
  constructor(listName: string, context?: WebPartContext, siteUrl?: string) {
    this.properties = new BaseListConfig();
    this.properties.listName = listName;

    if (context == undefined && siteUrl == undefined) {
      throw new Error("You must either supply a context or siteurl to use the PNPList Connector");
    }
    if (context != undefined) {
      sp.setup({
        spfxContext: context
      });
      this.web = sp.web;
    } else if (siteUrl != undefined) {
      this.web = new Web(siteUrl);
    }
  }

  /**
   * Set ListConfiguration Properties
   * @param listConfiguration
   */
  public setListConfiguration(listConfiguration: IListCfg) {
    this.properties = listConfiguration;
  }


  /**
   * Must be set to true if you want to use the list ensure function of this class
   * If no fieldconfiguration has been set before, this feature cannot be set to true
   * @param createOnFail
   */
  public setCreateOnFail(createOnFail: boolean): void {
    if (createOnFail == true && (this.properties.fields == undefined || this.properties.fields == null)) {
      throw new Error("You cannot enable createOnFail when no Field Configuration is defined. Use setListConfiguration to provide a Field Configuration");
    }
    this.createOnFail = createOnFail;
  }

  /**
   *  Get the Web Object
   */
  public getWeb(): Web {
    return this.web;
  }

  public getList(): List {
    return this.web.lists.getByTitle(this.properties.listName);
  }
  public addItem(newItem: any): Promise<ItemAddResult> {
    return this.addIem(newItem);
  }
  /**
   * @deprecated
   * @param newItem
   */
  public addIem(newItem: any): Promise<ItemAddResult> {
    if (this.createOnFail) {
      return this.ensureList().then((list: List): Promise<ItemAddResult> => {
        return list.items.add(newItem);
      });
    } else {
      //this.web.lists.ensure(this.properties.listName).then();
      let list = this.web.lists.getByTitle(this.properties.listName);
      return list.items.add(newItem);
    }
  }

  /**
   *
   * @param id list item id
   * @param item
   */

  public updateItem(id: number, item): Promise<ItemUpdateResult> {
    let list = this.web.lists.getByTitle(this.properties.listName);
    return list.items.getById(id).update(item);
  }

  /**
   * Deprecated because of Backwards compatibility - Use getItems
   * @param listTitle
   * @param query
   */

  public getItemByCAML(listTitle: string, query: CamlQuery): Promise<any> {
    if (this.createOnFail) {
      return this.ensureList().then((list: List): Promise<any> => {
        return list.getItemsByCAMLQuery(query);
      });
    } else {
      return this.web.lists.getByTitle(listTitle).getItemsByCAMLQuery(query).then((result: any[]) => {
        return result;
      }).catch((error: any) => {
        console.error(error);
        return error;
      });
    }
  }

  /**
   *
   * @param listTitle
   * @param selects
   */
  private getItems(listTitle: string, ...selects: string[]): Promise<any[]> {
    if (this.createOnFail) {
      return this.ensureList().then((list: List): Promise<any[]> => {
        return list.items.select(...selects).get();
      });
    } else {
      return this.web.lists.getByTitle(listTitle).select(...selects).items.get().then((items: any) => {
        return items;
      });
    }
  }

  /**
   *
   * @param data
   */
  public deleteItem(data): Promise<any> {
    if (this.createOnFail) {
      return this.ensureList().then((list: List): Promise<any> => {
        return list.items.getById(data.Id).delete().then((result) => {
          return Promise.resolve(true);
        }).catch((error) => {
          return Promise.reject(false);
        });
      }).catch((reject) => {
        console.error(reject);
        Promise.reject("Item cannot be deleted.");
      });
    }else{
      return this.web.lists.getByTitle(this.properties.listName).items.getById(data.Id).delete().then((result) => {
        console.log(result);
        return Promise.resolve(result);
      }).catch((error) => {
        console.error(error);
        return Promise.reject(error);
      });
    }
  }

  public addView(list: List, configuration: IViewConfig): Promise<any> {
    return list.views.getByTitle(configuration.ViewName).get().then((resolvedView) => {
      return Promise.resolve(resolvedView);
    }).catch((error) => {
      //fail silent because view dosent exist and we need the undefined value
    }).then((view) => {
      if (view != undefined) {
        return view;
      }

      return list.views.add(configuration.ViewName, false, { DefaultView: configuration.DefaultView, RowLimit: configuration.RowLimit, ViewQuery: configuration.ViewQuery }).then((va: ViewAddResult) => {

        return va.view.fields.removeAll().then(_ => {

          const batch = sp.web.createBatch();
          va.view.fields.inBatch(batch).removeAll();
          configuration.Fields.forEach(fieldName => {
            va.view.fields.inBatch(batch).add(fieldName);
            console.log(fieldName + " added");
          });
          return batch.execute().then((exec) => {
            console.log('Done');
            return Promise.resolve(view);
          }).catch((e) => {
            return Promise.reject(view);
          }
          );
          //return Promise.resolve(view);
        });
      }).catch((e) => {
        console.log(e);
      });
    });

  }

  /**
   * ensures that the list is available
   */
  public ensureList(): Promise<List> {
    return new Promise<List>((resolve: (list: List) => void, reject: (err: string) => void): void => {
      let listEnsureResults: ListEnsureResult;
      let fields: IListFieldCfg[] = this.properties.fields;
      //pnp.sp.web.lists.ensure(this.properties.listName).then(
      this.web.lists.ensure(this.properties.listName).then(
        (ler: ListEnsureResult) => {
          listEnsureResults = ler;
          if (!ler.created) {
            // resolve main promise
            resolve(ler.list);
            // break promise chain
            return Promise.reject(LIST_EXISTS);
          }
          //Create if not exists
          //const batch = pnp.sp.web.createBatch();
          const batch = this.web.createBatch();
          for (var i: number = 0; i < fields.length; i++) {
            let value = fields[i];
            let created = ler.list.fields.inBatch(batch).add(value.title, value.fieldType, value.fieldProperties).then((success) => {
              console.log("Success: " + value.title, success);
            }).catch((reason: any) => {
              console.error("Failure: " + value.title, reason.data);
            });
          }
          console.log(batch.execute());
        }
      ).then((): void => {
        // All of the items have been added within the batch
        resolve(listEnsureResults.list);
      }).catch((e: any): void => {
        if (e !== LIST_EXISTS) {
          reject(e);
        }
      });
    });
  }

}


