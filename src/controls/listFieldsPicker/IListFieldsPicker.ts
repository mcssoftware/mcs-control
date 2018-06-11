import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { IFieldOption } from "./IFieldOptions.";
import { ISPField } from "../../common/SPEntities";

export interface IListFieldsPickerProps {
    context: WebPartContext | ApplicationCustomizerContext;
    listTitle: string;
    className?: string;
    disabled?: boolean;
    selectedFields?: ISPField[];
    includeOrdering?: boolean;
    label?: string;
    placeHolder?: string;
    onSelectionChanged?: (newValue: ISPField[]) => void;
}

export interface IListFieldsPickerState {
    options: IFieldOption[];
    loading: boolean;
    selectedFields?: ISPField[];
}