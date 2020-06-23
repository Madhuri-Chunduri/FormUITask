import * as React from 'react';
import styles from './FormUi.module.scss';
import { spEventsParser } from "sharepoint-events-parser";
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { IFormUiProps } from './IFormUiProps';
import { escape } from '@microsoft/sp-lodash-subset';
import 'office-ui-fabric-react/dist/css/fabric.css';
import "./FormUi.sass";
import DragAndDrop from "./DragAndDropComponent";
import BootstrapModalComponent from "./BootstrapModalComponent";
import { Icon } from "office-ui-fabric-react/lib/Icon";
import { initializeIcons } from '@uifabric/icons';
import "office-ui-fabric-react";
import { sp } from "sp-pnp-js";
import * as pnp from "sp-pnp-js";
import { PeoplePicker } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { TaxonomyPicker, IPickerTerms } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";
import { IHttpClientOptions, HttpClientResponse, HttpClient } from '@microsoft/sp-http';
initializeIcons();

export default class FormUi extends React.Component<IFormUiProps, any> {
  private readonly attachmentsRef: React.RefObject<HTMLInputElement>

  constructor(props) {
    super(props);
    this.state = {
      title: "", comments: "", attachments: [],
      errors: { title: "", people: "", terms: "", comments: "" }, terms: [], people: [],
      showMandatoryFields: false, validationMessage: "", showFlyOut: false, showModalPopUp: false,
      showCommentsBox: false, redirectionUrl: escape(this.props.redirectionUrl), selectedFile: File
    }
    this.handleChange = this.handleChange.bind(this);
    this.submitPage = this.submitPage.bind(this);
    this.closePage = this.closePage.bind(this);
    this.getPeoplePickerItems = this.getPeoplePickerItems.bind(this);
    this.onTaxPickerChange = this.onTaxPickerChange.bind(this);
    this.keepBothFiles = this.keepBothFiles.bind(this);
    this.replaceFile = this.replaceFile.bind(this);
    this.setModalShow = this.setModalShow.bind(this);
    this.sendEmail = this.sendEmail.bind(this);
    this.attachmentsRef = React.createRef();
  }

  async componentDidMount() {
    console.log("Entered");
    let events = await sp.web.lists.getByTitle('SampleEvents').items.get();
    console.log("A : ", events);
    // const GetUserDetails = async (spHttpClient: SPHttpClient) => {
    //   return await this.context.spHttpClient
    //     .get(
    //       `https://technovert2020.sharepoint.com/sites/Technovert/_api/web/lists/getbytitle('SampleEvents')/items$select=*,Duration,RecurrenceData`,
    //       SPHttpClient.configurations.v1
    //     )
    //     .then((response: SPHttpClientResponse): any => {
    //       return response.json().then((responseJSON: any) => {
    //         console.log(responseJSON);
    //       });
    //     })
    // }

    console.log("Context : ", this.context);
    let eventDetails = this.props.context.spHttpClient.get(`https://technovert2020.sharepoint.com/sites/Technovert/_api/lists/GetByTitle('SampleEvents')/items`,
      SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json().then((responseJSON: any) => {
          console.log(responseJSON);
          let parsedArray = spEventsParser.parseEvents(responseJSON, 0, 10);
          console.log("Parsed Array : ", parsedArray);
          return responseJSON;
        });
      });

    console.log("Event Details : ", eventDetails);
    //   sp.web.lists.getByTitle('SampleEvents')
    //     .renderListDataAsStream({
    //       OverrideViewXml: `
    //           <QueryOptions>
    //               <ExpandRecurrence>TRUE</ExpandRecurrence>
    //           </QueryOptions>
    //       `
    //     })
    //     .then(console.log)
    //     .catch(console.log);
    //   const xml = `const xml = <View> \ <Query> \ <Where> \ <DateRangesOverlap> \ <FieldRef Name=EventDate /> \ <FieldRef Name=EndDate /> \ <FieldRef Name=RecurrenceID /> \ <Value Type=DateTime> \ <Today /> \ </Value> \ </DateRangesOverlap> \ </Where> \ <OrderBy> \ <FieldRef Name=EventDate /> \ </OrderBy>\ </Query> \ <RowLimit>10</RowLimit> \ </View>;`;
    //   sp.web.lists.getByTitle('SampleEvents')
    //     .renderListDataAsStream({
    //       OverrideViewXml: xml
    //     })
    //     .then(console.log)
    //     .catch(console.log);
    //   console.log("Item details : ", GetUserDetails);
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
    else if (this.state.people.length > 1) {
      count += 1;
      errors.people = "Requestors cannot be more than one";
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
      this.setState({
        validationMessage: "* Please fill the below fields with valid data",
        errors: errors,
      });
      return false;
    }

    return true;
  }

  async submitPage() {
    if (!this.validateForm()) {
      pnp.setup({
        spfxContext: this.props.context
      });
      const multiTermNoteFieldName = 'TermField_0';

      let termsString: string = '';
      this.state.terms.forEach(term => {
        termsString += `-1;#${term.name}|${term.key};#`;
      });

      const folderItem = await sp.web.rootFolder.folders.getByName("CrawlLibrary").folders.get().then(async result => {
        console.log("Folders : ", result);
        await sp.web.rootFolder.folders.getByName("CrawlLibrary").folders.getByName(result[1].Name).files.get().then(result => {
          console.log("Items : ", result);
        })
        const insertedFile = await sp.web.rootFolder.folders.getByName("CrawlLibrary").folders.getByName(result[1].Name).files.add(
          this.state.attachments[0].name,
          this.state.attachments[0],
          true
        ).then(result => {
          console.log("Items : ", result);
          return result;
        })
        const item = insertedFile.file.getItem();
        (await item).update({
          DescriptionText: this.state.comments
        }).then(console.log);
        return result;

      });

      const spfxList = sp.web.lists.getByTitle('FormList');
      spfxList.getListItemEntityTypeFullName()
        .then((entityTypeFullName) => {
          spfxList.fields.getByTitle(multiTermNoteFieldName).get()
            .then((taxNoteField) => {
              const multiTermNoteField = taxNoteField.InternalName;
              const updateObject = {
                Title: this.state.title,
                RequestorId:
                  this.state.people[0]["id"]
                ,
                Comments: this.state.comments
              };
              updateObject[multiTermNoteField] = termsString;

              spfxList.items.add(updateObject, entityTypeFullName)
                .then(
                  async (result) => {
                    result.item.attachmentFiles.addMultiple(this.state.attachments).then(result => {
                      window.location.href = this.state.redirectionUrl;
                    });
                  }
                )
            });
        });

    }
    else {
      this.setState({ validationMessage: "* Please fill all the mandatory fields", showMandatoryFields: true })
    }
  }

  closePage() {
    window.location.href = this.state.redirectionUrl;
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
    this.setState({ attachments: attachments, showCommentsBox: showCommentsBox });
  }

  replaceFile(file) {
    let attachments = this.state.attachments;
    let attachmentNames = []
    attachments.forEach(file => {
      attachmentNames.push(file.name);
    });
    var index = attachmentNames.indexOf(file.name);
    attachments[index] = {
      name: file.name,
      content: file
    }
    this.setState({ attachments: attachments, showModalPopUp: false });
  }

  keepBothFiles(file) {
    let attachments = this.state.attachments;
    let attachmentNames = []
    attachments.forEach(file => {
      attachmentNames.push(file.name);
    });
    var count = this.getCount(file, attachmentNames);
    var typeIndex = file.name.lastIndexOf(".");         //if count>0, then a duplicate file already exists in currently uploaded files[i]
    var fileName = file.name.substring(0, typeIndex) + "(" + count + ")" + file.name.substring(typeIndex);
    attachments.push({
      name: fileName,
      content: file
    });
    this.setState({ attachments: attachments, showModalPopUp: false });
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
          showModalPopUp = true;
        }
      }
      else showModalPopUp = true;
      this.setState({ attachments: attachments, showCommentsBox: true, showModalPopUp: showModalPopUp, selectedFile: files[i] });
    }
  }


  onTaxPickerChange(terms: IPickerTerms) {
    let presentTerms = this.state.terms;
    presentTerms.push(terms);
    let errors = this.state.errors;
    if (presentTerms.length > 0) {
      errors.terms = "";
    }
    else errors.terms = "Mandatory Field";
    this.setState({ terms: terms, errors: errors });
  }

  getPeoplePickerItems(items: any[]) {
    let errors = this.state.errors;
    if (items.length == 1) {
      errors.people = "";
    }
    else if (items.length > 1) errors.people = "Requestors cannot be more than one";
    else errors.people = "Mandatory Field";
    this.setState({ people: items, errors: errors });
  }

  setModalShow() {
    this.setState({ showModalPopUp: false });
  }

  sendEmail(): Promise<HttpClientResponse> {
    const postURL = "https://prod-09.centralindia.logic.azure.com:443/workflows/979cef66a55a433da50ffdd2f9a19577/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=3SlYq_5IADCEMu1owf1wfB5PXnLaG0q6fFKFNZUCVjQ";

    const body: string = JSON.stringify({
      'emailaddress': "madhuri.c@technovert.net",
      'emailSubject': "Test mail",
      'emailBody': "This mail is actiavted from Power Automate Flow",
    });

    const requestHeaders: Headers = new Headers();
    requestHeaders.append('Content-type', 'application/json');

    const httpClientOptions: IHttpClientOptions = {
      body: body,
      headers: requestHeaders
    };

    console.log("Sending Email");

    return this.props.context.httpClient.post(
      postURL,
      HttpClient.configurations.v1,
      httpClientOptions)
      .then((response: HttpClientResponse): Promise<HttpClientResponse> => {
        console.log("Email sent.");
        return response.json();
      });
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
            <p>Title U*</p>
          </div>
        </div>
        <div className="ms-Grid-row">
          <div><input type="text" name="title" value={this.state.title} className={errors.title.length > 0 && this.state.showMandatoryFields ? 'emptyTextField' : 'textField'} onChange={this.handleChange} /></div>
          <div className={`ms-Grid-col ms-sm11 ms-md11 ms-lg11 ms-xl11  ${(errors.title.length > 0) ? "errorTitle" : "hideElement"}`}>{errors.title}</div>
        </div>

        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm11 ms-md11 ms-lg11 ms-xl11">
            <PeoplePicker
              context={this.props.context}
              titleText={"Requestor *"}
              personSelectionLimit={3}
              groupName=""
              showtooltip={false}
              ensureUser={true}
              disabled={false}
              selectedItems={this.getPeoplePickerItems} />
            <p className={(errors.people.length > 0) ? "error" : "hideElement"}>{errors.people}</p>
          </div>
        </div>

        <div className="ms-Grid-row">
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
        <div className="ms-Grid-row attachments">
          <div className="commentsField">
            <div><p className="commentsTitle"> Comments </p></div>
            <div><textarea value={this.state.comments} className="commentsArea" name="comments" onChange={this.handleChange} /></div>
            <div><p className={(errors.comments.length > 0) ? "error" : "hideElement"}>{errors.comments}</p></div>
          </div>
          <div className="attachmentsField">
            <div className="attachmentsTitle">
              <input ref={this.attachmentsRef} type="file" style={{ display: "none" }} onChange={this.onInsertFile.bind(this)} />
              <p> Attachments
            <button className="addAttachmentsButton" onClick={() => this.attachmentsRef.current.click()}> <Icon iconName="Attach" /> Add file </button></p>
              <div className="descriptionBox">
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
          </div>
        </div>
        <div className="ms-Grid-row">
          <input type="button" className="ms-Grid-col ms-sm2 ms-md2 ms-lg2 ms-xl2 submitButton" value="Submit" onClick={this.submitPage} />
          <input type="button" className="ms-Grid-col ms-sm2 ms-md2 ms-lg2 ms-xl2 submitButton" value="Close" onClick={this.closePage} />
          <input type="button" className="ms-Grid-col ms-sm2 ms-md2 ms-lg2 ms-xl2 submitButton" value="Send Email" onClick={this.sendEmail} />
        </div>
        {/* {this.state.showModalPopUp ? <ModalComponent isOpen={this.state.showModalPopUp} file={this.state.selectedFile} keepBothFiles={this.keepBothFiles} replaceFile={this.replaceFile} /> : ""} */}
        {/* {this.state.showModalPopUp ? <ColorPickerDialogContent message="Success" close={() => { }} /> : ""} */}
        {this.state.showModalPopUp ? <BootstrapModalComponent show={this.state.showModalPopUp}
          onHide={this.setModalShow} file={this.state.selectedFile} keepBothFiles={this.keepBothFiles} replaceFile={this.replaceFile} /> : ""}
      </div >
    );
  }
}
