import * as React from 'react';
import styles from './FormUi.module.scss';
import { IFormUiProps } from './IFormUiProps';
import { escape } from '@microsoft/sp-lodash-subset';
import 'office-ui-fabric-react/dist/css/fabric.css';
import "./FormUi.sass";
import { sp } from "sp-pnp-js";
import * as pnp from "sp-pnp-js";
import DragAndDrop from "./DragAndDropComponent";
import TaxonomyPickerComponent from "./TaxonomyPickerComponent";
import ModalComponent from "./ModalComponent";
import { Icon } from "office-ui-fabric-react/lib/Icon";
import { initializeIcons } from '@uifabric/icons';
import { PeoplePicker } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { TaxonomyPicker, IPickerTerms } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";
initializeIcons();

export default class FormUi extends React.Component<IFormUiProps, any> {
    private readonly attachmentsRef: React.RefObject<HTMLInputElement>

    constructor(props) {
        super(props);
        this.state = {
            title: "", comments: "", attachments: [],
            errors: { title: "", people: "", terms: "", comments: "" }, terms: [], people: [],
            showMandatoryFields: false, validationMessage: "", showFlyOut: false, showModalPopUp: false,
            showCommentsBox: false, redirectionUrl: escape(this.props.redirectionUrl)
        }
        this.handleChange = this.handleChange.bind(this);
        this.submitPage = this.submitPage.bind(this);
        this.openFlyOut = this.openFlyOut.bind(this);
        this.getPeoplePickerItems = this.getPeoplePickerItems.bind(this);
        this.onTaxPickerChange = this.onTaxPickerChange.bind(this);
        this.attachmentsRef = React.createRef();
    }

    handleChange = (event) => {
        const target = event.target;
        const fieldName = target.name;
        let errors = this.state.errors;
        if (event.target.value.length > 0) {
            if (fieldName == "title") errors.title = "";
            if (fieldName == "comments") errors.comments = "";
        }
        else if (fieldName == "title") errors.title = "Mandatory Field";
        this.setState({ [fieldName]: event.target.value, errors: errors });
    }

    validateField(fieldName, fieldValue) {
        let errors = this.state.errors;
        switch (fieldName) {
            case "title": if (fieldValue.length == 0) {
                errors.title = "* Title cannot be empty";
            }
            else errors.title = "";
                break;

            case "requestor": if (fieldValue.length == 0) {
                errors.requestor = "* Requestor cannot be empty";
            }
            else errors.requestor = "";
                break;

            case "office": if (fieldValue.length == 0) {
                errors.office = "* Office cannot be empty";
            }
            else errors.office = "";
                break;
        }
        this.setState({ errors: errors });
    }

    validateForm() {
        let count = 0;
        let errors = this.state.errors;

        if (this.state.terms.length == 0) {
            errors.terms = "Mandatory Field";
            count += 1;
        }
        else { errors.terms = "" }

        if (this.state.people.length == 0) {
            errors.people = "Mandatory Field";
            count += 1;
        }
        else { errors.people = "" }

        if (this.state.title.length == 0) {
            errors.title = "Mandatory Field";
            count += 1;
        }
        else { errors.title = "" }

        if (this.state.attachments.length > 0) {
            if (this.state.comments.length == 0) {
                count += 1;
                errors.comments = "Mandatory Field";
            }
            else errors.comments = "";
        }
        else errors.comments = "";

        if (count > 0) {
            this.setState({ errors: errors });
            return false;
        }

        return true;
    }

    submitPage() {
        if (this.validateForm()) {
            pnp.setup({
                spfxContext: this.props.context
            });
            return sp.web.lists
                .getByTitle("FormList")
                .items.add({
                    "Title": this.state.title,
                    "RequestorId": this.state.people[0],
                    "Comments": this.state.description
                }).then(
                    async (result) => {
                        let attachments = []
                        this.state.files.forEach(file => {
                            attachments.push({
                                name: file.name,
                                content: file
                            })
                        }
                        );
                        result.item.attachmentFiles.addMultiple(this.state.attachments).then(result => {
                            window.location.href = this.state.redirectionUrl;
                        });
                    }
                )

        }
        else {
            this.setState({ validationMessage: "* Please fill all the mandatory fields", showMandatoryFields: true })
        }
    }

    onInsertFile(event) {
        event.stopPropagation();
        event.preventDefault();
        var file = event.target.files[0];
        let files = [];
        files.push(file);
        event.target.value = null;
        this.handleDrop(files);
    }

    removeFile(file) {
        let attachments = this.state.attachments;
        let showCommentsBox = this.state.showCommentsBox;
        if (attachments.length == 1) {
            showCommentsBox = false;
        }
        let index = attachments.indexOf(file);
        attachments.splice(index, 1);
        // this.renameFiles(file);
        this.setState({ attachments: attachments, showCommentsBox: showCommentsBox });
    }

    renameFiles(file) {
        let fileName = file.name;
        let index = fileName.lastIndexOf("(");
        let endIndex = fileName.lastIndexOf(")");
        let typeIndex = fileName.lastIndexOf(".");
        let type = fileName.substring(typeIndex);
        let attachments = this.state.attachments;
        let actualFileName = file.name;
        let count = 0;
        if (index != -1) {
            count = parseInt(fileName.substring(index + 1, endIndex));
            actualFileName = fileName.substring(0, index) + fileName.substring(typeIndex);
            //file.name = fileName.substring(0, typeIndex) + "(" + changedCount + ")" + fileName.substring(typeIndex);
            console.log("Actual File name : ", actualFileName);
        }
        for (var i = 0; i < attachments.length; i++) {
            let name = attachments[i].name;
            index = name.lastIndexOf("(");
            endIndex = name.lastIndexOf(")");
            typeIndex = name.lastIndexOf(".");
            type = name.substring(typeIndex);
            let currentFileName = name.substring(0, index) + name.substring(typeIndex);
            console.log("Current file name : ", currentFileName);
            if (currentFileName == actualFileName) {
                let currentFileCount = parseInt(name.substring(index + 1, endIndex));
                console.log("Current file count : ", currentFileCount);
                if (count < currentFileCount) {
                    currentFileCount = currentFileCount - 1;
                    if (currentFileCount > 0) attachments[i].name = name.substring(0, index) + "(" + currentFileCount + ")" + name.substring(typeIndex);
                    else attachments[i].name = name.substring(0, index) + name.substring(typeIndex);
                }
            }
        }
        this.setState({ attachments: attachments });
    }

    getCount(givenFile: { name: string; }, existingFiles: any[]) {
        var typeIndex, count = 1, substringName = "";
        existingFiles.forEach(file => {
            let index = file.lastIndexOf("(");
            let endIndex = file.lastIndexOf(")");
            typeIndex = file.lastIndexOf(".");
            type = file.substring(typeIndex);
            if (index != -1) {
                substringName = file.substring(0, index);
                typeIndex = file.lastIndexOf(".");
                var type = file.substring(typeIndex);
                var currentFile = substringName + type;
                if (currentFile == givenFile.name) {
                    var currentCount = parseInt(file.substring(index + 1, endIndex));
                    currentCount += 1;
                    if (currentCount > count) {
                        count = currentCount;
                    }
                }
            }
        });
        return count;
    }


    handleDrop = (files) => {
        console.log("Files ", files);
        let attachments = this.state.attachments;
        let showModalPopUp = this.state.showModalPopUp;
        let attachmentNames = []
        attachments.forEach(file => {
            attachmentNames.push(file.name);
        });
        for (var i = 0; i < files.length; i++) {
            var count = this.getCount(files[i], attachmentNames); //Returns the count of occurences of file
            if (count == 1) {
                let index = attachmentNames.indexOf(files[i].name);
                if (index == -1) { //means no file with same name exists in currently uploaded files[i], so we check for it's existence in past files[i]
                    attachments.push({
                        name: files[i].name,
                        content: files[i]
                    });
                }
                else { //if index!=-1 and count=1, then there exists only a single file ex: Technovert.png
                    // var typeIndex = files[i].name.lastIndexOf(".");
                    // var fileName = files[i].name.substring(0, typeIndex) + "(" + count + ")" + files[i].name.substring(typeIndex);
                    // attachments.push({
                    //   name: fileName,
                    //   content: files[i]
                    // });
                    showModalPopUp = true;
                }
            }
            // else {
            //   typeIndex = files[i].name.lastIndexOf(".");         //if count>0, then a duplicate file already exists in currently uploaded files[i]
            //   var fileName = files[i].name.substring(0, typeIndex) + "(" + count + ")" + files[i].name.substring(typeIndex);
            //   attachments.push({
            //     name: fileName,
            //     content: files[i]
            //   });
            // }
        }
        this.setState({ attachments: attachments, showCommentsBox: true, showModalPopUp: true });
    }

    openFlyOut() {
        const element: React.ReactElement<IFormUiProps> = React.createElement(
            TaxonomyPickerComponent,
            {
                redirectionUrl: this.props.redirectionUrl,
                context: this.context
            }
        );

        this.setState({ showFlyOut: true });
    }

    onTaxPickerChange(terms: IPickerTerms) {
        let presentTerms = this.state.terms;
        presentTerms.push(terms);
        let errors = this.state.errors;
        if (presentTerms.length > 0) {
            errors.terms = "";
        }
        else errors.terms = "Mandatory Field";
        this.setState({ terms: presentTerms, errors: errors });
    }

    getPeoplePickerItems(items: any[]) {
        let presentPeople = this.state.people;
        presentPeople.push(items);
        let errors = this.state.errors;
        if (presentPeople.length > 0) {
            errors.people = "";
        }
        else errors.people = "Mandatory Field";

        this.setState({ people: presentPeople, errors: errors });
    }

    toggleModal() {
        this.setState({ showModalPopUp: false });
    }

    public render() {
        let errors = this.state.errors;

        return (
            <div className="ms-Grid formBody" dir="ltr">
                {this.state.validationMessage.length > 0 ? <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12 ms-xl4 error">
                        <p>{this.state.validationMessage}</p>
                    </div>
                </div> : ""}

                <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12 ms-xl4">
                        <p>Title *</p>
                    </div>
                </div>
                <div className="ms-Grid-row">
                    <div><input type="text" name="title" value={this.state.title} className={`ms-Grid-col ms-sm11 ms-md11 ms-lg11 ms-xl11 ${errors.title.length > 0 && this.state.showMandatoryFields ? 'emptyTextField' : 'textField'}`} onChange={this.handleChange} /></div>
                    <div className={`ms-Grid-col ms-sm11 ms-md11 ms-lg11 ms-xl11  ${(errors.title.length > 0) ? "errorTitle" : "hideElement"}`}>{errors.title}</div>
                </div>

                <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm11 ms-md11 ms-lg11 ms-xl11">
                        {/* <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12 ms-xl12">
            <p>Requestor {errors.requestor.length > 0 ? <span className="error">{errors.requestor}</span> : ""}</p>
          </div>
        </div>
        <div className="ms-Grid-row">
          <input className={`ms-Grid-col ms-sm11 ms-md11 ms-lg11 ms-xl11 ${this.state.showMandatoryFields && errors.requestor.length > 0 ? 'emptyTextField' : 'textField'}`} type="text" name="requestor" value={this.state.requestor} onChange={this.handleChange} /> */}
                        <PeoplePicker
                            context={this.props.context}
                            titleText={"Requestor *"}
                            personSelectionLimit={3}
                            groupName=""
                            showtooltip={true}
                            disabled={false}
                            selectedItems={this.getPeoplePickerItems} />
                        <p className={(errors.people.length > 0) ? "error" : "hideElement"}>{errors.people}</p>
                    </div>
                </div>

                <div className="ms-Grid-row">
                    {/* <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12 ms-xl4">
            <p>Office {errors.office.length > 0 ? <span className="error">{errors.office}</span> : ""}</p>
          </div>
        </div>
        <div className="ms-Grid-row"> */}
                    {/* <input type="text" name="office" value={this.state.office} className={`ms-Grid-col ms-sm10 ms-md10 ms-lg10 ms-xl10 ${this.state.showMandatoryFields && errors.requestor.length > 0 ? 'emptyTextField' : 'textField'}`} onChange={this.handleChange} /> */}
                    <div className="ms-Grid-col ms-sm11 ms-md11 ms-lg11 ms-xl11 termIcon">
                        <TaxonomyPicker allowMultipleSelections={true}
                            termsetNameOrID="JobRoles"
                            panelTitle="Select Term"
                            label="Term Field *"
                            context={this.props.context}
                            onChange={this.onTaxPickerChange}
                            isTermSetSelectable={false}
                        />
                        <p className={(errors.terms.length > 0) ? "error" : "hideElement"}>{errors.terms}</p>

                    </div>
                </div>
                <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm5 ms-md5 ms-lg5 ms-xl5">
                        <p className="commentsTitle"> Comments </p>
                    </div>
                    <div className="ms-Grid-col ms-sm5 ms-md5 ms-lg5 ms-xl5 attachmentsField">
                        <div className="ms-Grid-row">
                            <div className="ms-Grid-col ms-sm4 ms-md4 ms-lg4 ms-xl4"> <p> Attachments </p> </div>
                            <input ref={this.attachmentsRef} type="file" style={{ display: "none" }} onChange={this.onInsertFile.bind(this)} />
                            <div className="ms-Grid-col ms-sm5 ms-md5 ms-lg5 ms-xl5">
                                <button className="addAttachmentsButton" onClick={() => this.attachmentsRef.current.click()}> <Icon iconName="Attach" /> Add file </button>
                            </div>
                        </div>
                    </div>
                </div>

                <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm5 ms-md5 ms-lg5 ms-xl5">
                        <div><textarea value={this.state.comments} className="commentsArea" name="comments" onChange={this.handleChange} /></div>
                        <div><p className={(errors.comments.length > 0) ? "error" : "hideElement"}>{errors.comments}</p></div>
                    </div>
                    <div className="ms-Grid-col ms-sm5 ms-md5 ms-lg5 ms-xl5 descriptionBox">
                        <DragAndDrop handleDrop={this.handleDrop}>
                            <div>{this.state.attachments.map((file) =>
                                <div>{file.name}
                                    <input className="removeFileButton" type="button" value="x" onClick={() => this.removeFile(file)} />
                                </div>
                            )}
                            </div>
                        </DragAndDrop>
                    </div>
                </div>

                <div className="ms-Grid-row">
                    <input type="button" className="ms-Grid-col ms-sm2 ms-md2 ms-lg2 ms-xl2 submitButton" value="Submit" onClick={this.submitPage} />
                    <input type="button" className="ms-Grid-col ms-sm2 ms-md2 ms-lg2 ms-xl2 submitButton" value="Close" onClick={this.submitPage} />
                </div>
                {this.state.showModalPopUp ? <ModalComponent isOpen={this.state.showModalPopUp} /> : ""}
            </div >
        );
    }
}
