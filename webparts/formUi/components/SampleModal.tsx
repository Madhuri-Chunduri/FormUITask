import * as React from 'react';
import {
    ColorPicker,
    PrimaryButton,
    Button,
    DialogFooter,
    DialogContent,
    IColor
} from 'office-ui-fabric-react';

interface IColorPickerDialogContentProps {
    message: string;
    close: () => void;
}

class ColorPickerDialogContent extends React.Component<IColorPickerDialogContentProps, {}> {
    private _pickedColor: IColor;

    constructor(props) {
        super(props);
        // Default Color
        this._pickedColor = props.defaultColor || { hex: 'FFFFFF', str: '', r: null, g: null, b: null, h: null, s: null, v: null };
    }

    public render(): JSX.Element {
        return <DialogContent
            title='Color Picker'
            subText={this.props.message}
            onDismiss={this.props.close}
            showCloseButton={true}
        >
            <ColorPicker color={this._pickedColor} onChange={this._onColorChange} />
            <DialogFooter>
                <Button text='Cancel' title='Cancel' onClick={this.props.close} />
            </DialogFooter>
        </DialogContent>;
    }

    private _onColorChange = (ev: React.SyntheticEvent<HTMLElement, Event>, color: IColor) => {
        this._pickedColor = color;
    }
}

export default ColorPickerDialogContent;