import * as React from "react";
import {
    DetailsList,
    DetailsListLayoutMode,
    Selection,
    SelectionMode,
    IColumn,
    IGroup,
    TextField,
    autobind,
} from "office-ui-fabric-react";
import styles from "./ListView.module.scss";
import { IListViewProps, IListViewState, IViewField, IGrouping, GroupOrder } from "./IListView";
import { findIndex, has, sortBy, isEqual, cloneDeep } from "@microsoft/sp-lodash-subset";
import { FileTypeIcon, IconType } from "../fileTypeIcon/index";
import { IGroupsItems } from "./IListView";

// tslint:disable:typedef
/**
 * File type icon component
 */
export class ListView extends React.Component<IListViewProps, IListViewState> {
    private _selection: Selection;

    constructor(props: IListViewProps) {
        super(props);

        // initialize state
        this.state = {
            items: [],
        };

        // binding the functions
        this._columnClick = this._columnClick.bind(this);

        if (typeof this.props.selection !== "undefined" && this.props.selection !== null) {
            // initialize the selection
            this._selection = new Selection({
                // create the event handler when a selection changes
                onSelectionChanged: () => this.props.selection(this._selection.getSelection()),
            });
        }
    }

    /**
     * Life cycle hook when component is mounted
     */
    public componentDidMount(): void {
        this._processProperties();
    }

    /**
     * Life cycle hook when component did update after state or property changes
     * @param prevProps
     * @param prevState
     */
    public componentDidUpdate(prevProps: IListViewProps, prevState: IListViewState): void {
        // select default items
        this._setSelectedItems();

        if (!isEqual(prevProps, this.props)) {
            this._processProperties();
        }
    }

    /**
     * Select all the items that should be selected by default
     */
    private _setSelectedItems(): void {
        if (this.props.items &&
            this.props.items.length > 0 &&
            this.props.defaultSelection &&
            this.props.defaultSelection.length > 0) {
            for (const index of this.props.defaultSelection) {
                if (index > -1) {
                    this._selection.setIndexSelected(index, true, false);
                }
            }
        }
    }

    /**
     * Specify result grouping for the list rendering
     * @param items
     * @param groupByFields
     */
    private _getGroups(items: any[], groupByFields: IGrouping[], level: number = 0, startIndex: number = 0): IGroupsItems {
        // group array which stores the configured grouping
        const groups: IGroup[] = [];
        const updatedItemsOrder: any[] = [];
        // check if there are group by fields set
        if (groupByFields) {
            const groupField: IGrouping = groupByFields[level];
            // check if grouping is configured
            if (groupByFields && groupByFields.length > 0) {
                // create grouped items object
                const groupedItems: any = {};
                items.forEach((item: any) => {
                    let groupName: any = item[groupField.name];
                    // check if the group name exists
                    if (typeof groupName === "undefined") {
                        // set the default empty label for the field
                        groupName = "Empty Group Label";
                    }
                    // check if group name is a number, this can cause sorting issues
                    if (typeof groupName === "number") {
                        groupName = `${groupName}.`;
                    }

                    // check if current group already exists
                    if (typeof groupedItems[groupName] === "undefined") {
                        // create a new group of items
                        groupedItems[groupName] = [];
                    }
                    groupedItems[groupName].push(item);
                });

                // sort the grouped items object by its key
                const sortedGroups: any = {};
                let groupNames: string[] = Object.keys(groupedItems);
                groupNames = groupField.order === GroupOrder.ascending ? groupNames.sort() : groupNames.sort().reverse();
                groupNames.forEach((key: string) => {
                    sortedGroups[key] = groupedItems[key];
                });

                // loop over all the groups
                // tslint:disable-next-line:forin
                for (const groupItems in sortedGroups) {
                    // retrieve the total number of items per group
                    const totalItems: number = groupedItems[groupItems].length;
                    // create the new group
                    const group: IGroup = {
                        name: groupItems === "undefined" ? "Empty Group Label" : groupItems,
                        key: groupItems === "undefined" ? "Empty Group Label" : groupItems,
                        startIndex,
                        count: totalItems,
                    };
                    // check if child grouping available
                    if (groupByFields[level + 1]) {
                        // get the child groups
                        const subGroup = this._getGroups(groupedItems[groupItems], groupByFields, (level + 1), startIndex);
                        subGroup.items.forEach((item) => {
                            updatedItemsOrder.push(item);
                        });
                        group.children = subGroup.groups;
                    } else {
                        // add the items to the updated items order array
                        groupedItems[groupItems].forEach((item) => {
                            updatedItemsOrder.push(item);
                        });
                    }
                    // increase the start index for the next group
                    startIndex = startIndex + totalItems;
                    groups.push(group);
                }
            }
        }
        return {
            items: updatedItemsOrder,
            groups,
        };
    }

    /**
     * Process all the component properties
     */
    private _processProperties() {
        const tempState: IListViewState = cloneDeep(this.state);
        let columns: IColumn[] = null;
        // check if a set of items was provided
        if (typeof this.props.items !== "undefined" && this.props.items !== null) {
            tempState.items = this._flattenItems(this.props.items);
        }

        // check if an icon needs to be shown
        if (typeof this.props.iconFieldName !== "undefined" && this.props.iconFieldName !== null) {
            if (columns === null) { columns = []; }
            const iconColumn = this._createIconColumn(this.props.iconFieldName);
            columns.push(iconColumn);
        }

        // check if view fields were provided
        if (typeof this.props.viewFields !== "undefined" && this.props.viewFields !== null) {
            if (columns === null) { columns = []; }
            columns = this._createColumns(this.props.viewFields, columns);
        }

        // add the columns to the temporary state
        tempState.columns = columns;

        // add grouping to the list view
        const grouping = this._getGroups(tempState.items, this.props.groupByFields);
        if (grouping.groups.length > 0) {
            tempState.groups = grouping.groups;
            // update the items
            tempState.items = grouping.items;
        } else {
            tempState.groups = null;
        }

        // update the current component state with the new values
        this.setState(tempState);
    }

    /**
     * Flatten all objects in every item
     * @param items
     */
    private _flattenItems(items: any[]): any[] {
        // flatten items
        const flattenItems = items.map((item) => {
            // flatten all objects in the item
            return this._flattenItem(item);
        });
        return flattenItems;
    }

    /**
     * Flatten all object in the item
     * @param item
     */
    private _flattenItem(item: any): any {
        const flatItem = {};
        for (const parentPropName in item) {
            // check if property already exists
            if (!item.hasOwnProperty(parentPropName)) {
                continue;
            }

            // check if the property is of type object
            if ((typeof item[parentPropName]) === "object") {
                // flatten every object
                const flatObject = this._flattenItem(item[parentPropName]);
                for (const childPropName in flatObject) {
                    if (!flatObject.hasOwnProperty(childPropName)) {
                        continue;
                    }
                    flatItem[`${parentPropName}.${childPropName}`] = flatObject[childPropName];
                }
            } else {
                flatItem[parentPropName] = item[parentPropName];
            }
        }
        return flatItem;
    }

    /**
     * Create an icon column rendering
     * @param iconField
     */
    private _createIconColumn(iconFieldName: string): IColumn {
        return {
            key: "fileType",
            name: "File Type",
            iconName: "Page",
            isIconOnly: true,
            fieldName: "fileType",
            minWidth: 16,
            maxWidth: 16,
            onRender: (item: any) => {
                return (
                    <FileTypeIcon type={IconType.image} path={item[iconFieldName]} />
                );
            },
        };
    }

    /**
     * Returns required set of columns for the list view
     * @param viewFields
     */
    private _createColumns(viewFields: IViewField[], currentColumns: IColumn[]): IColumn[] {
        const maxWidth: number = ((this.refs.outerContainer as HTMLElement).offsetWidth / (viewFields.length));
        viewFields.forEach((field) => {
            currentColumns.push({
                key: field.name,
                name: field.displayName || field.name,
                fieldName: field.name,
                minWidth: field.displayName.length * 6,
                //  field.minWidth || 50,
                maxWidth: field.maxWidth || maxWidth,
                onRender: this._fieldRender(field),
                onColumnClick: this._columnClick,
            });
        });
        return currentColumns;
    }

    /**
     * Check how field needs to be rendered
     * @param field
     */
    private _fieldRender(field: IViewField): any | void {
        // check if a render function is specified
        if (field.render) {
            return field.render;
        }
        // check if the URL property is specified
        // if (/LinkTitle/gi.test(field.name)) {
        //    return (item: any, index?: number, column?: IColumn) => {
        //        return <a href={item[field.linkPropertyName]}>{item[column.fieldName]}</a>;
        //    };
        // }
    }

    /**
     * Check if sorting needs to be set to the column
     * @param ev
     * @param column
     */
    private _columnClick(ev: React.MouseEvent<HTMLElement>, column: IColumn): void {
        // find the field in the viewFields list
        const columnIdx = findIndex(this.props.viewFields, (field) => field.name === column.key);
        // check if the field has been found
        if (columnIdx !== -1) {
            const field = this.props.viewFields[columnIdx];
            // check if the field needs to be sorted
            if (has(field, "sorting")) {
                // check if the sorting option is true
                if (field.sorting) {
                    const sortDescending = typeof column.isSortedDescending === "undefined" ? false : !column.isSortedDescending;
                    const sortedItems = this._sortItems(this.state.items, column.key, sortDescending);
                    // update the columns
                    const sortedColumns = this.state.columns.map((c) => {
                        if (c.key === column.key) {
                            c.isSortedDescending = sortDescending;
                            c.isSorted = true;
                        } else {
                            c.isSorted = false;
                            c.isSortedDescending = false;
                        }
                        return c;
                    });
                    // update the grouping
                    const groupedItems = this._getGroups(sortedItems, this.props.groupByFields);
                    // update the items and columns
                    this.setState({
                        items: groupedItems.groups.length > 0 ? groupedItems.items : sortedItems,
                        columns: sortedColumns,
                        groups: groupedItems.groups.length > 0 ? groupedItems.groups : null,
                    });
                }
            }
        }
    }

    /**
     * Sort the list of items by the clicked column
     * @param items
     * @param columnName
     * @param descending
     */
    private _sortItems(items: any[], columnName: string, descending = false): any[] {
        const ascItems = sortBy(items, [columnName]);
        return descending ? ascItems.reverse() : ascItems;
    }

    /**
     * Default React component render method
     */
    // tslint:disable-next-line:member-ordering
    public render(): React.ReactElement<IListViewProps> {
        return (
            <div ref="outerContainer" className={styles.listview}>
                {this.props.showFilter &&
                    <div className={styles.row}>
                        <div className={styles.searchColumn}>
                            <TextField
                                className={styles.searchBox}
                                label="Search"
                                onKeyPress={this._onKeyPress}
                                onBeforeChange={this._onFilterChanged} />
                        </div>
                    </div>
                }
                <DetailsList
                    className={this._getHeightCss()}
                    items={this._getFilterItems()}
                    columns={this.state.columns}
                    groups={this.state.groups}
                    selectionMode={SelectionMode.single}
                    selectionPreservedOnEmptyClick={true}
                    selection={this._selection}
                    layoutMode={DetailsListLayoutMode.justified}
                    compact={this.props.compact}
                    setKey="ListViewControl" />
            </div>
        );
    }

    @autobind
    private _onKeyPress(e: any): void {
        if (e.charCode === 13) {
            e.preventDefault();
        }
    }

    private _getHeightCss(): string {
        if (typeof this.props.heightCss !== "string" && (this.props.heightCss as string).length < 1) {
            return "";
        }
        const exp: RegExp = new RegExp(this.props.heightCss, "gi");
        return exp.test(styles.viewPort30) ? styles.viewPort30 : exp.test(styles.viewPort60) ? styles.viewPort60 : styles.viewPort90;
    }

    private _getFilterItems(): any[] {
        const { filterText } = this.state;

        if (this.props.showFilter && typeof filterText === "string" && filterText.length > 2) {
            const { columns, items } = this.state;
            const filterRegex: RegExp = new RegExp(this.state.filterText, "gi");
            return items.filter((item) => {
                let canDisplayItem: boolean = false;
                // tslint:disable-next-line:prefer-for-of
                for (let i: number = 0; i < columns.length; i++) {
                    const key: string = columns[i].key;
                    const value: any = item[key];
                    if (typeof value === "string" && value.length > 0 && filterRegex.test(value)) {
                        canDisplayItem = true;
                        break;
                    }
                }
                return canDisplayItem;
            });
        }
        return this.state.items;
    }

    @autobind
    private _onFilterChanged(text: string) {
        this.setState({
            ...this.state,
            filterText: text,
        });
    }
}
