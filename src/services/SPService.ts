import { ISPService, ILibsOptions, LibsOrderBy } from "./ISPService";
import { ISPLists, ISPField } from "../common/SPEntities";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { Web } from "sp-pnp-js";
export default class SPService implements ISPService {

  constructor(private _context: WebPartContext | ApplicationCustomizerContext) { }

  /**
   * Get lists or libraries
   * @param options
   */
  public getLibs(options?: ILibsOptions): Promise<ISPLists> {
    let filter: string = "";
    const listPropsSelects: string[] = ["Title", "id", "BaseTemplate"];

    if (options.baseTemplate) {
      filter += `BaseTemplate eq ${options.baseTemplate}`;
    }

    if (options.includeHidden === false) {
      filter = filter + " and Hidden eq false";
    }

    return new Web(this._context.pageContext.web.absoluteUrl).lists
      .select(...listPropsSelects)
      .filter(filter)
      .orderBy(options.orderBy === LibsOrderBy.Id ? "Id" : "Title").get();
  }

  public getFields(listTitle: string): Promise<ISPField[]> {

    const select: string[] = ["InternalName", "Sortable", "Title", "TypeAsString", "IsDependentLookup", "LookupField", "PrimaryFieldId", "Id",
      "LookupList", "DependentLookupInternalNames"];
    const url: string = "https://mcssoftwaresolutions.sharepoint.com/sites/LmsDev2/Legismng2018"; // this._context.pageContext.web.absoluteUrl
    return (new Web(url)).lists.getByTitle(listTitle).fields.filter("Hidden eq false")
      .select(...select)
      .get();
  }
}
