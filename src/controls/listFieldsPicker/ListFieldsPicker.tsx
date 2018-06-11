import React = require("react");
import { IListFieldsPickerProps, IListFieldsPickerState } from "./IListFieldsPicker";
import { IFieldOption } from "./IFieldOptions.";
import { SPServiceFactory } from "../../services/SPServiceFactory";
import { ISPService } from "../../services/ISPService";
import { ISPField } from "../../common/SPEntities";
import { findIndex, clone, find } from "@microsoft/sp-lodash-subset";
import { IDropdownProps, Dropdown, Spinner, SpinnerSize, autobind, List, Label, css, Checkbox, IDropdownOption } from "office-ui-fabric-react";
import styles from "./listFieldsPicker.module.scss";

export class ListFieldsPicker extends React.Component<IListFieldsPickerProps, IListFieldsPickerState> {
    private _listFields: ISPField[];

    constructor(props: IListFieldsPickerProps) {
        super(props);
        this._listFields = [];
        this.state = {
            options: [],
            loading: false,
        };
    }

    public render(): JSX.Element {
        const { loading, options } = this.state;
        const { className, disabled, includeOrdering, label, placeHolder } = this.props;

        if (includeOrdering === true) {
            return (
                <div className={styles.listFieldsPicker}>
                    {loading && <Spinner className={styles.spinner} size={SpinnerSize.xSmall} />}
                    {!loading && <div>
                        <Label>{label}</Label>
                        <List
                            items={options}
                            onRenderCell={this._onRenderCell} />
                    </div>}
                </div>
            );
        } else {
            const dropdownOptions: IDropdownProps = {
                className,
                options,
                disabled: (loading || disabled),
                label,
                placeHolder,
                onChanged: this._onChanged,
                multiSelect: true,
            };
            return (
                <div className={styles.listFieldsPicker}>
                    {loading && <Spinner className={styles.spinner} size={SpinnerSize.xSmall} />}
                    {!loading && <Dropdown {...dropdownOptions} />}
                </div>
            );
        }
    }

    public componentDidMount(): void {
        this._loadFields();
    }

    private _loadFields(): void {
        const { context, listTitle, selectedFields } = this.props;
        // show the loading indicator and disable the dropdown
        this.setState({ loading: true });

        const service: ISPService = SPServiceFactory.createService(context, true, 5000);
        service.getFields(listTitle).then((results) => {
            if (selectedFields) {
                results = results.sort((a, b) => {
                    let aindex: number = findIndex(selectedFields, (f) => f.InternalName === a.InternalName);
                    if (aindex < 0) {
                        aindex = 9999;
                    }
                    let bindex: number = findIndex(selectedFields, (f) => f.InternalName === b.InternalName);
                    if (bindex < 0) {
                        bindex = 9999;
                    }
                    if (aindex === bindex && aindex === 9999) {
                        if (a.Title < b.Title) {
                            return -1;
                        }
                        if (a.Title > b.Title) {
                            return 1;
                        }
                        return 0;
                    }
                    return aindex - bindex;
                });
            }
            this._listFields = results;
            const options: IFieldOption[] = [];
            // start mapping the lists to the dropdown option
            results.map((field, index) => {
                options.push({
                    key: field.InternalName,
                    text: field.Title,
                    orderIndex: index,
                    selected: selectedFields ? findIndex(selectedFields, (f) => f.InternalName === field.InternalName) >= 0 : false,
                });
            });

            // hide the loading indicator and set the dropdown options and enable the dropdown
            this.setState({
                loading: false,
                options,
                selectedFields: typeof selectedFields === "undefined" && selectedFields === null && Array.isArray(selectedFields) ? selectedFields : [],
            });
        });
    }

    @autobind
    private _onChanged(option: IFieldOption, index?: number): void {
        const { onSelectionChanged } = this.props;
        let { selectedFields } = this.state;
        if (typeof selectedFields === "undefined" && selectedFields === null && !Array.isArray(selectedFields)) {
            selectedFields = [];
        }
        if (option.selected) {
            selectedFields.push(this._listFields[index]);
        } else {
            selectedFields = selectedFields.filter((f) => f.InternalName === option.key);
        }
        if (onSelectionChanged) {
            onSelectionChanged(selectedFields);
        }
    }

    private _onRenderCell = (item: IFieldOption, index: number): JSX.Element => {
        const ddlOptions: IDropdownOption[] = this._listFields.map((e, i) => {
            return {
                key: i + 1,
                text: (i + 1).toString(),
            };
        });
        return (
            <div className="ms-itemCell" data-is-focusable={true}>
                <div className={css(styles.itemContent,
                    index % 2 === 0 && styles.itemContentEven,
                    index % 2 === 1 && styles.itemContentOdd,
                )}>
                    <Checkbox label={item.text} onChange={(ev: any, isChecked: boolean) => { this._onCheckboxChange(item, index, isChecked); }} />
                    <Dropdown options={ddlOptions}
                        className={styles.listDropdown}
                        disabled={this.props.disabled || false}
                        selectedKey={this._getSelectedKey(index)}
                        onChanged={(option) => { this._onlistDdlChanged(option, index); }}
                        multiSelect={false} />
                </div>
            </div>
        );
    }

    @autobind
    private _getSelectedKey(index: number): number {
        return this.state.options[index].orderIndex + 1;
    }

    @autobind
    private _onCheckboxChange(item: IFieldOption, index: number, isChecked: boolean): void {
        const { onSelectionChanged } = this.props;
        let { selectedFields } = this.state;
        if (isChecked) {
            selectedFields.push(this._listFields[index]);
            item.selected = true;
        } else {
            selectedFields = selectedFields.filter((f) => f.InternalName === item.key);
            item.selected = false;
        }
        if (onSelectionChanged) {
            onSelectionChanged(selectedFields);
        }
    }

    @autobind
    private _onlistDdlChanged(selectedOption: IFieldOption, index?: number): void {
        let { selectedFields } = this.state;
        const options: IFieldOption[] = clone(this.state.options);
        const { onSelectionChanged } = this.props;
        const newOrder: number = selectedOption.key as number - 1;
        const oldorder: number = options[index].orderIndex;
        if (newOrder === oldorder) {
            return;
        } else {
            if (newOrder < oldorder) {
                options.forEach((o, i) => {
                    if (o.orderIndex >= newOrder && i !== index && o.orderIndex < oldorder) {
                        o.orderIndex = o.orderIndex + 1;
                    }
                });
            }
            else { // if (newOrder > oldorder)
                options.forEach((o, i) => {
                    if (i !== index) {
                        if (o.orderIndex > oldorder && o.orderIndex <= newOrder) {
                            o.orderIndex = o.orderIndex - 1;
                        }
                    }
                });
            }
        }
        options[index].orderIndex = newOrder;
        selectedFields = options.filter((o) => o.selected).sort((a, b) => a.orderIndex - b.orderIndex).map((o) => find(this._listFields, (f) => f.InternalName === o.key));
        this.setState({ options, selectedFields });
        if (onSelectionChanged) {
            onSelectionChanged(selectedFields);
        }
    }
}
