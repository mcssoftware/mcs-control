import * as React from "react";
import { IDropdownOption, IDropdownProps, Dropdown } from "office-ui-fabric-react/lib/components/Dropdown";
import { Spinner, SpinnerSize } from "office-ui-fabric-react/lib/components/Spinner";
import { IListPickerProps, IListPickerState } from "./IListPicker";
import { ISPService } from "../../services/ISPService";
import { SPServiceFactory } from "../../services/SPServiceFactory";
import styles from "./ListPicker.module.scss";
import { autobind } from "office-ui-fabric-react";

// tslint:disable:jsdoc-format

/**
* Empty list value, to be checked for single list selection
*/
const EMPTY_LIST_KEY: string = "NO_LIST_SELECTED";

/**
* Renders the controls for the ListPicker component
*/
export class ListPicker extends React.Component<IListPickerProps, IListPickerState> {
  private _options: IDropdownOption[] = [];
  private _selectedList: string | string[];

  /**
  * Constructor method
  */
  constructor(props: IListPickerProps) {
    super(props);
    this.state = {
      options: this._options,
      loading: false,
    };
  }

  /**
  * Lifecycle hook when component is mounted
  */
  public componentDidMount(): void {
    this._loadLists();
  }

  /**
   * componentDidUpdate lifecycle hook
   * @param prevProps
   * @param prevState
   */
  public componentDidUpdate(prevProps: IListPickerProps, prevState: IListPickerState): void {
    if (
      prevProps.baseTemplate !== this.props.baseTemplate ||
      prevProps.includeHidden !== this.props.includeHidden ||
      prevProps.orderBy !== this.props.orderBy ||
      prevProps.selectedList !== this.props.selectedList
    ) {
      this._loadLists();
    }
  }

  /**
  * Renders the ListPicker controls with Office UI Fabric
  */
  public render(): JSX.Element {
    const { loading, options, selectedList } = this.state;
    const { className, disabled, multiSelect, label, placeHolder } = this.props;

    const dropdownOptions: IDropdownProps = {
      className,
      options,
      disabled: (loading || disabled),
      label,
      placeHolder,
      onChanged: this._onChanged,
    };

    if (multiSelect === true) {
      dropdownOptions.multiSelect = true;
      dropdownOptions.selectedKeys = selectedList as string[];
    } else {
      dropdownOptions.selectedKey = selectedList as string;
    }

    return (
      <div className={styles.listPicker}>
        {loading && <Spinner className={styles.spinner} size={SpinnerSize.xSmall} />}
        <Dropdown {...dropdownOptions} />
      </div>
    );
  }

  /**
 * Loads the list from SharePoint current web site
 */
  private _loadLists(): void {
    const { context, baseTemplate, includeHidden, orderBy, multiSelect, selectedList } = this.props;

    // show the loading indicator and disable the dropdown
    this.setState({ loading: true });

    const service: ISPService = SPServiceFactory.createService(context, true, 5000);
    service.getLibs({
      baseTemplate,
      includeHidden,
      orderBy,
    }).then((results) => {
      // start mapping the lists to the dropdown option
      results.value.map((list) => {
        this._options.push({
          key: list.Id,
          text: list.Title,
        });
      });

      if (multiSelect !== true) {
        // add option to unselct list
        this._options.unshift({
          key: EMPTY_LIST_KEY,
          text: "",
        });
      }

      this._selectedList = this.props.selectedList;

      // hide the loading indicator and set the dropdown options and enable the dropdown
      this.setState({
        loading: false,
        options: this._options,
        selectedList: this._selectedList,
      });
    });
  }

  /**
  * Raises when a list has been selected
  * @param option the new selection
  * @param index the index of the selection
  */
  @autobind
  private _onChanged(option: IDropdownOption, index?: number): void {
    const { multiSelect, onSelectionChanged } = this.props;

    if (multiSelect === true) {
      if (!this._selectedList) {
        this._selectedList = [] as string[];
      }

      const selectedLists: string[] = this._selectedList as string[];
      // check if option was selected
      if (option.selected) {
        selectedLists.push(option.key as string);
      } else {
        // filter out the unselected list
        this._selectedList = selectedLists.filter((list) => list !== option.key);
      }
    } else {
      this._selectedList = option.key as string;
    }

    if (onSelectionChanged) {
      onSelectionChanged(this._selectedList);
    }
  }
}
