import ReactModal from "react-modal";
import * as React from "react";
import styles from "./ModalComponent.module.scss";
import { DialogContent, PrimaryButton } from "office-ui-fabric-react";
import "./ModalComponent.sass";

class ModalComponent extends React.Component<any, any> {
    constructor(props) {
        super(props);
        this.state = {
            showModal: this.props.isOpen
        };

        this.handleOpenModal = this.handleOpenModal.bind(this);
        this.handleCloseModal = this.handleCloseModal.bind(this);
        this.keepBoth = this.keepBoth.bind(this);
        this.replaceFile = this.replaceFile.bind(this);
    }

    handleOpenModal() {
        this.setState({ showModal: true });
    }

    handleCloseModal() {
        this.setState({ showModal: false });
    }

    replaceFile() {
        const replaceFile = this.props.replaceFile;
        replaceFile(this.props.file);
    }

    keepBoth() {
        const keepBothFiles = this.props.keepBothFiles;
        keepBothFiles(this.props.file);
    }

    componentWillReceiveProps() {
        console.log("Entered Will Receive");
        this.setState({ showModal: true });
    }

    render() {
        console.log("Props : ", this.props.isOpen);
        console.log("State : ", this.state.showModal);
        return (
            <div className="ms-Grid" dir="ltr">
                <DialogContent className="ms-Grid-row">
                    <ReactModal className="ms-Grid-col ms-sm5 ms-md5 ms-lg5 ms-xl2 modalBody"
                        isOpen={this.state.showModal}
                        contentLabel="Minimal Modal Example"
                    >
                        <h3>File already exists</h3>
                        <div><PrimaryButton className="primaryButton" onClick={this.handleCloseModal}>Disard New File</PrimaryButton></div>
                        <div><PrimaryButton className="primaryButton" onClick={this.replaceFile}>Replace Old File</PrimaryButton></div>
                        <div><PrimaryButton className="primaryButton" onClick={this.keepBoth}>Keep Both</PrimaryButton></div>
                    </ReactModal>
                </DialogContent>
            </div>
        );
    }
}

export default ModalComponent;