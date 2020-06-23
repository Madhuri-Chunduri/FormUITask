import * as React from "react";
import { TaxonomyPicker, IPickerTerms } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";
import { IFormUiProps } from './IFormUiProps';
import styles from './FormUi.module.scss';

export default class PnPTaxonomyPicker extends React.Component<IFormUiProps, any> {
    private onTaxPickerChange(terms: IPickerTerms) {
        console.log("Terms", terms);
    }

    render(): React.ReactElement<IFormUiProps> {
        console.log("Taxonomy Picker Context : ", this.props.context);
        return (
            <div className={styles.pnpTaxonomyPicker}>
                <div className={styles.container}>
                    <div className={styles.row}>
                        <div className={styles.column}>
                            <TaxonomyPicker allowMultipleSelections={true}
                                termsetNameOrID="JobRoles"
                                panelTitle="Select Term"
                                label=""
                                context={this.props.context}
                                onChange={this.onTaxPickerChange}
                                isTermSetSelectable={false} />
                        </div>
                    </div>
                </div>
            </div>
        );
    }
}  