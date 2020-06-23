import * as React from "react";
import { Modal, Button } from "react-bootstrap";
import 'bootstrap/dist/css/bootstrap.min.css';
import "./ModalComponent.sass";
import { PrimaryButton } from "office-ui-fabric-react";
require("react-bootstrap/ModalHeader");
require("react-bootstrap/ModalBody")
require("react-bootstrap/ModalFooter")

class BootstrapModalComponent extends React.Component<any, any>{

    constructor(props) {
        super(props);
        this.replaceFile = this.replaceFile.bind(this);
        this.keepBoth = this.keepBoth.bind(this);
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
        return (
            // <Modal.Dialog>
            //     <Modal.Header closeButton>
            //         <Modal.Title>Modal title</Modal.Title>
            //     </Modal.Header>

            //     <Modal.Body>
            //         <p>Modal body text goes here.</p>
            //     </Modal.Body>

            //     <Modal.Footer>
            //         <Button variant="secondary">Close</Button>
            //         <Button variant="primary">Save changes</Button>
            //     </Modal.Footer>
            // </Modal.Dialog>
            <Modal
                {...this.props}
                size="lg"
                aria-labelledby="contained-modal-title-vcenter"
                centered
            >

                <Modal.Body className="bootstrapModalBody">
                    <h4>File already exists!!</h4>
                    <div><PrimaryButton className="primaryButton" onClick={this.props.onHide}>Disard New File</PrimaryButton></div>
                    <div><PrimaryButton className="primaryButton" onClick={this.replaceFile}>Replace Old File</PrimaryButton></div>
                    <div><PrimaryButton className="primaryButton" onClick={this.keepBoth}>Keep Both</PrimaryButton></div>
                </Modal.Body>

            </Modal>

        )
    }
}

export default BootstrapModalComponent