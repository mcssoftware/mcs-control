import * as React from "react";
import styles from "./McsControlTest.module.scss";
import { IMcsControlTestProps } from "./IMcsControlTestProps";
import { ListFieldsPicker } from "../../../controls/listFieldsPicker";

export default class McsControlTest extends React.Component<IMcsControlTestProps, {}> {
  public render(): React.ReactElement<IMcsControlTestProps> {
    return (
      <div className={styles.mcsControlTest}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <ListFieldsPicker context={this.context}
                listTitle="Bills"
                disabled={false}
                includeOrdering={true}
                label="Select fields" />
            </div>
          </div>
        </div>
      </div>
    );
  }
}
