import { ISPLists, ISPField } from "../common/SPEntities";

export enum LibsOrderBy {
    Id = 1,
    Title,
}
/**
 * Options used to sort and filter
 */
export interface ILibsOptions {
    orderBy?: LibsOrderBy;
    baseTemplate?: number;
    includeHidden?: boolean;
}
export interface ISPService {
    /**
     * Get the lists from SharePoint
     * @param options Options used to order and filter during the API query
     */
    getLibs(options?: ILibsOptions): Promise<ISPLists>;
    getFields(listTitle: string): Promise<ISPField[]>;
}