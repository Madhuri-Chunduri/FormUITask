import * as React from "react";
import * as pnp from "sp-pnp-js";
import { sp } from "sp-pnp-js";
import { IStackTokens, Stack } from 'office-ui-fabric-react/lib/Stack';
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';

export interface IReactGetItemsState {
    items: IDropdownOption[];
}

const stackTokens: IStackTokens = { childrenGap: 20 };
const dropdownStyles: Partial<IDropdownStyles> = {
    dropdown: { width: 300 }
};

class DropDownComponent extends React.Component<any, any>{
    constructor(props) {
        super(props);
        this.state = {
            items: []
        };
    }

    async componentDidMount() {
        let items = [];
        await sp.web.lists.getByTitle("AddressBook").items.select('Title').get().then(function (data) {
            for (var k in data) {
                items.push({ key: data[k].Title, text: data[k].Title });
            }
        });
        this.setState({ items });
        console.log(items);
        return items;
    }

    render() {
        return (
            <Stack tokens={stackTokens}>
                <Dropdown placeholder="Select an option" label="Contact Name" options={this.state.items} styles={dropdownStyles} />
            </Stack>
        )
    }

}

export default DropDownComponent;